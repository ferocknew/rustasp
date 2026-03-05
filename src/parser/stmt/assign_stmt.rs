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
            match self.peek_ahead(pos) {
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
                _ => return false,
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

        // 解析后缀（索引访问或属性访问）
        loop {
            match self.peek() {
                // 索引访问: arr(index)
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
                    expr = if args.len() == 1 {
                        Expr::Index {
                            object: Box::new(expr),
                            index: Box::new(args.into_iter().next().unwrap()),
                        }
                    } else {
                        return Err(ParseError::ParserError(
                            "Invalid index in assignment target".to_string(),
                        ));
                    };
                }

                // 属性访问: obj.property
                Token::Dot => {
                    self.advance();
                    let prop = self.expect_ident()?;
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
    pub fn parse_expr_stmt(&mut self) -> Result<Option<Stmt>, ParseError> {
        let expr = self.parse_expr(0)?;
        self.skip_newlines();
        Ok(Some(Stmt::Expr(expr)))
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
