//! 赋值语句和表达式语句解析

use crate::ast::{Expr, Stmt};
use crate::parser::lexer::Token;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析标识符开头的语句（赋值或表达式）
    ///
    /// VBScript 中 `x = 1` 是赋值语句，但 `if x = 1 then` 中的 `x = 1` 是比较表达式。
    /// 我们通过 lookahead 判断：如果看到 `identifier =`，则解析为赋值。
    pub fn parse_ident_stmt(&mut self) -> Result<Option<Stmt>, ParseError> {
        // Lookahead: 检查是否是赋值语句
        // 赋值形式: x = value 或 x(index) = value 或 x.property = value
        if self.is_assignment_start() {
            return self.parse_assignment();
        }

        // 否则是表达式语句（比较表达式）
        let expr = self.parse_expr(0)?;
        self.skip_newlines();
        Ok(Some(Stmt::Expr(expr)))
    }

    /// 检查是否是赋值语句的开始
    ///
    /// 赋值语句的模式：
    /// - identifier = ...  (变量赋值)
    /// - identifier(...) = ...  (索引赋值)
    /// - identifier.property = ...  (属性赋值)
    fn is_assignment_start(&self) -> bool {
        // 必须从标识符开始
        if !matches!(self.peek(), Token::Ident(_)) {
            return false;
        }

        // Lookahead 寻找赋值符号
        let mut pos = 1;
        loop {
            let tok = self.peek_ahead(pos);
            match tok {
                // 直接赋值: x = ...
                Token::Eq => {
                    // 检查 = 后面是否是新的一行或语句结束
                    // 如果是，则是比较表达式；否则是赋值
                    // 在 VBScript 中，单独一行的 x = 1 总是赋值
                    return true;
                }

                // 可能是索引访问，继续查找
                Token::LParen => {
                    pos += 1;
                    // 跳过括号内容
                    let mut depth = 1;
                    while depth > 0 && pos < 100 {  // 防止无限循环
                        match self.peek_ahead(pos) {
                            Token::LParen => {
                                depth += 1;
                                pos += 1;
                            }
                            Token::RParen => {
                                depth -= 1;
                                pos += 1;
                            }
                            Token::Eof => return false,
                            _ => pos += 1,
                        }
                    }
                    // 继续检查是否有 =
                }

                // 可能是属性访问，继续查找
                Token::Dot => {
                    pos += 1;
                    // 跳过属性名
                    match self.peek_ahead(pos) {
                        Token::Ident(_) => pos += 1,
                        _ => return false,
                    }
                }

                // 其他情况不是赋值
                _ => {
                    return false;
                }
            }
        }
    }

    /// 解析赋值语句
    ///
    /// 手动解析左侧表达式，避免把 = 当作比较运算符
    fn parse_assignment(&mut self) -> Result<Option<Stmt>, ParseError> {
        // 手动解析左侧（变量、索引访问或属性访问）
        let target = self.parse_lhs()?;

        // 期望 =
        if !self.check(&Token::Eq) {
            return Err(ParseError::ParserError(format!(
                "Expected '=' in assignment, got {:?}",
                self.peek()
            )));
        }
        self.advance(); // 消耗 =

        // 解析右侧（完整表达式）
        let value = self.parse_expr(0)?;

        // 跳过冒号（语句分隔符）
        self.match_token(&Token::Colon);
        self.skip_newlines();

        Ok(Some(Stmt::Assignment { target, value }))
    }

    /// 解析赋值左侧表达式（不包含 = 作为比较运算符）
    fn parse_lhs(&mut self) -> Result<Expr, ParseError> {
        // 解析标识符
        let name = self.expect_ident()?;
        let mut expr = Expr::Variable(name);

        // 解析后缀（索引访问、属性访问）
        loop {
            match self.peek() {
                // 索引访问: arr(index) - 这是赋值目标，不是函数调用
                Token::LParen => {
                    self.advance();
                    let mut args = vec![];
                    if !self.check(&Token::RParen) {
                        args.push(self.parse_expr(0)?);
                        while self.match_token(&Token::Comma) {
                            args.push(self.parse_expr(0)?);
                        }
                    }
                    self.expect(Token::RParen)?;

                    // 在赋值左侧，括号总是表示索引访问
                    // 支持多维数组: arr(0,0) 被解析为嵌套索引 ((arr(0))(0))
                    for arg in args {
                        expr = Expr::Index {
                            object: Box::new(expr),
                            index: Box::new(arg),
                        };
                    }
                }

                // 属性访问: obj.property
                Token::Dot => {
                    self.advance();
                    // 点号后面允许标识符或关键字作为属性名
                    let prop = match self.peek().clone() {
                        Token::Ident(name) => {
                            self.advance();
                            name
                        }
                        Token::Keyword(kw) => {
                            self.advance();
                            kw.as_str().to_string()
                        }
                        _ => {
                            return Err(ParseError::ParserError(format!(
                                "Expected identifier after '.', got {:?}",
                                self.peek()
                            )))
                        }
                    };
                    expr = Expr::Property {
                        object: Box::new(expr),
                        property: prop,
                    };
                }

                // 不是后缀，结束
                _ => break,
            }
        }

        Ok(expr)
    }

    /// 解析表达式语句
    ///
    /// 支持 VBScript 风格的方法调用：Response.Write "Hello"
    /// 不需要括号，直接在属性名后跟参数
    pub fn parse_expr_stmt(&mut self) -> Result<Option<Stmt>, ParseError> {
        let expr = self.parse_expr(0)?;

        // 检查是否是 obj.property 形式，后面是否有参数
        // VBScript 支持 Response.Write "Hello" 这种无括号调用
        let stmt = if let Expr::Property { object, property } = expr {
            // 检查下一个 token 是否可能是参数
            // 如果是字符串、数字、标识符等，则收集参数
            if self.is_argument_start() {
                let mut args = vec![];
                // 收集所有参数（用空格分隔）
                while self.is_argument_start() {
                    args.push(self.parse_expr(0)?);
                    // VBScript 中参数用空格分隔，不是逗号
                    // 但我们也兼容逗号
                    if !self.match_token(&Token::Comma) {
                        // 检查是否还有更多参数
                        if !self.is_argument_start() {
                            break;
                        }
                    }
                }
                // 转换为方法调用
                Stmt::Expr(Expr::Method {
                    object,
                    method: property,
                    args,
                })
            } else {
                // 不是方法调用，保持原样
                Stmt::Expr(Expr::Property { object, property })
            }
        } else {
            Stmt::Expr(expr)
        };

        self.skip_newlines();
        Ok(Some(stmt))
    }

    /// 检查是否是参数的开始
    fn is_argument_start(&self) -> bool {
        match self.peek() {
            Token::String(_) => true,
            Token::Number(_) => true,
            Token::Ident(_) => true,
            Token::LParen => true,
            _ => false,
        }
    }

    /// 解析 Set 语句
    pub fn parse_set(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(crate::parser::keyword::Keyword::Set)?;
        let target = self.parse_lhs()?;
        self.expect(Token::Eq)?;
        let value = self.parse_expr(0)?;

        // 跳过冒号（语句分隔符）
        self.match_token(&Token::Colon);
        self.skip_newlines();

        Ok(Some(Stmt::Set { target, value }))
    }
}
