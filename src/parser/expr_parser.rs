//! 极简 Pratt Parser - 表达式解析器
//!
//! 使用绑定优先级（binding power）实现，无泛型嵌套，编译快速

use super::lexer::Token;
use super::ParseError;
use crate::ast::{BinaryOp, Expr, UnaryOp};

/// 表达式解析器
pub struct ExprParser {
    tokens: Vec<Token>,
    pos: usize,
}

impl ExprParser {
    /// 创建新的表达式解析器
    pub fn new(tokens: Vec<Token>) -> Self {
        ExprParser { tokens, pos: 0 }
    }

    /// 解析表达式（入口函数）
    pub fn parse(&mut self) -> Result<Expr, ParseError> {
        let expr = self.parse_expression(0)?;
        self.expect_end()?;
        Ok(expr)
    }

    /// 核心解析函数 - Pratt 算法
    ///
    /// min_bp: 最小绑定优先级，低于此优先级的运算符会停止解析
    fn parse_expression(&mut self, min_bp: u8) -> Result<Expr, ParseError> {
        // 1. 解析前缀表达式（左侧）
        let mut lhs = self.parse_prefix()?;

        // 2. 循环处理中缀运算符
        loop {
            // 查看下一个 token 是否是中缀运算符
            let (l_bp, r_bp) = match self.infix_binding_power() {
                Some(bp) => bp,
                None => break,
            };

            // 如果左侧优先级小于要求的最小优先级，停止
            if l_bp < min_bp {
                break;
            }

            // 消耗运算符并转换为 BinaryOp（避免借用问题）
            let op_token = self.next()?.clone();
            let op = self.token_to_binary_op(&op_token)?;

            // 解析右侧表达式（使用右侧优先级）
            let rhs = self.parse_expression(r_bp)?;

            // 构建二元运算 AST
            lhs = Expr::Binary {
                left: Box::new(lhs),
                op,
                right: Box::new(rhs),
            };
        }

        Ok(lhs)
    }

    /// 解析前缀表达式
    fn parse_prefix(&mut self) -> Result<Expr, ParseError> {
        // 获取并克隆 token，避免借用冲突
        let token = self.peek().clone();

        match token {
            // 字面量
            Token::Number(n) => {
                self.next()?;
                Ok(Expr::Number(n))
            }
            Token::String(s) => {
                self.next()?;
                Ok(Expr::String(s))
            }
            Token::Boolean(b) => {
                self.next()?;
                Ok(Expr::Boolean(b))
            }

            // 变量或标识符
            Token::Ident(name) => {
                self.next()?;
                self.parse_postfix(Expr::Variable(name))
            }

            // 一元运算符
            Token::Minus => {
                self.next()?;
                let rhs = self.parse_expression(self.unary_binding_power())?;
                Ok(Expr::Unary {
                    op: UnaryOp::Neg,
                    operand: Box::new(rhs),
                })
            }

            Token::Keyword(kw) if kw.is_unary_op() => {
                self.next()?;
                let rhs = self.parse_expression(self.unary_binding_power())?;
                Ok(Expr::Unary {
                    op: UnaryOp::Not,
                    operand: Box::new(rhs),
                })
            }

            // 括号表达式
            Token::LParen => {
                self.next()?;
                let expr = self.parse_expression(0)?;
                self.expect(Token::RParen)?;
                // 括号表达式后也可能有后缀，如 (expr)(args)
                self.parse_postfix(expr)
            }

            _ => Err(ParseError::ParserError(format!(
                "Unexpected token in expression: {:?}",
                token
            ))),
        }
    }

    /// 解析后缀表达式（成员访问、方法调用、索引访问）
    ///
    /// 支持：
    /// - obj.property
    /// - obj.method(args)
    /// - obj(index)
    /// - 链式访问：obj.method()(index).property
    fn parse_postfix(&mut self, mut lhs: Expr) -> Result<Expr, ParseError> {
        loop {
            match self.peek() {
                // 成员访问：obj.property 或 obj.method(...)
                Token::Dot => {
                    self.next()?; // 消耗 .
                    let name = self.expect_ident()?;
                    lhs = Expr::Property {
                        object: Box::new(lhs),
                        property: name,
                    };
                }

                // 方法调用或索引访问：(args)
                Token::LParen => {
                    self.next()?; // 消耗 (
                    let mut args = Vec::new();
                    if !self.is_at(Token::RParen) {
                        loop {
                            args.push(self.parse_expression(0)?);
                            if !self.match_comma() {
                                break;
                            }
                        }
                    }
                    self.expect(Token::RParen)?;

                    // 如果 lhs 是 Property，转换为 Method
                    // 否则作为 Index 处理
                    lhs = match lhs {
                        Expr::Property { object, property } => Expr::Method {
                            object,
                            method: property,
                            args,
                        },
                        _ => {
                            // 单参数时作为索引，多参数报错
                            if args.len() == 1 {
                                Expr::Index {
                                    object: Box::new(lhs),
                                    index: Box::new(args.into_iter().next().unwrap()),
                                }
                            } else if args.is_empty() {
                                // 空参数：如果是变量，当作无参函数调用
                                // 否则返回空索引（保持原样）
                                match lhs {
                                    Expr::Variable(name) => Expr::Call {
                                        name,
                                        args: vec![],
                                    },
                                    _ => Expr::Index {
                                        object: Box::new(lhs),
                                        index: Box::new(Expr::Empty),
                                    },
                                }
                            } else {
                                return Err(ParseError::ParserError(
                                    "Multiple indices not supported, use array".to_string(),
                                ));
                            }
                        }
                    };
                }

                // 不再是后缀操作符，结束循环
                _ => break,
            }
        }

        Ok(lhs)
    }

    /// 匹配逗号，返回是否匹配
    fn match_comma(&mut self) -> bool {
        if let Some(Token::Comma) = self.peek_if() {
            let _ = self.next();
            true
        } else {
            false
        }
    }

    /// 获取中缀运算符的绑定优先级
    ///
    /// 返回 (left_bp, right_bp)
    /// - left_bp: 左侧结合强度
    /// - right_bp: 右侧结合强度
    fn infix_binding_power(&self) -> Option<(u8, u8)> {
        match self.peek() {
            // 逻辑或 (最低优先级)
            Token::Keyword(kw) if kw.is_or() => Some((1, 2)),

            // 逻辑与
            Token::Keyword(kw) if kw.is_and() => Some((3, 4)),

            // 比较（相等性）
            Token::Eq | Token::Ne => Some((5, 6)),

            // 比较（大小）
            Token::Lt | Token::Le | Token::Gt | Token::Ge => Some((7, 8)),

            // 字符串连接
            Token::Ampersand => Some((9, 10)),

            // 加减
            Token::Plus | Token::Minus => Some((11, 12)),

            // 乘除
            Token::Star | Token::Slash | Token::Backslash => Some((13, 14)),

            // 幂运算（右结合）
            Token::Caret => Some((16, 15)), // 注意：right_bp < left_bp 表示右结合

            _ => None,
        }
    }

    /// 一元运算符的绑定优先级
    fn unary_binding_power(&self) -> u8 {
        17 // 高于所有二元运算符
    }

    /// 将 Token 转换为 BinaryOp
    fn token_to_binary_op(&self, token: &Token) -> Result<BinaryOp, ParseError> {
        match token {
            Token::Plus => Ok(BinaryOp::Add),
            Token::Minus => Ok(BinaryOp::Sub),
            Token::Star => Ok(BinaryOp::Mul),
            Token::Slash => Ok(BinaryOp::Div),
            Token::Backslash => Ok(BinaryOp::IntDiv),
            Token::Caret => Ok(BinaryOp::Pow),
            Token::Ampersand => Ok(BinaryOp::Concat),
            Token::Eq => Ok(BinaryOp::Eq),
            Token::Ne => Ok(BinaryOp::Ne),
            Token::Lt => Ok(BinaryOp::Lt),
            Token::Le => Ok(BinaryOp::Le),
            Token::Gt => Ok(BinaryOp::Gt),
            Token::Ge => Ok(BinaryOp::Ge),
            Token::Keyword(kw) if kw.is_and() => Ok(BinaryOp::And),
            Token::Keyword(kw) if kw.is_or() => Ok(BinaryOp::Or),
            _ => Err(ParseError::ParserError(format!(
                "Expected binary operator, got {:?}",
                token
            ))),
        }
    }

    // ========== 辅助方法 ==========

    fn peek(&self) -> &Token {
        if self.pos < self.tokens.len() {
            &self.tokens[self.pos]
        } else {
            &Token::Eof
        }
    }

    fn peek_if(&self) -> Option<&Token> {
        if self.pos < self.tokens.len() {
            Some(&self.tokens[self.pos])
        } else {
            None
        }
    }

    fn is_at(&self, token: Token) -> bool {
        std::mem::discriminant(self.peek()) == std::mem::discriminant(&token)
    }

    fn next(&mut self) -> Result<&Token, ParseError> {
        if self.pos < self.tokens.len() {
            let token = &self.tokens[self.pos];
            self.pos += 1;
            Ok(token)
        } else {
            Err(ParseError::ParserError("Unexpected end of input".to_string()))
        }
    }

    fn expect(&mut self, expected: Token) -> Result<&Token, ParseError> {
        let token = self.next()?;
        if std::mem::discriminant(token) == std::mem::discriminant(&expected) {
            Ok(token)
        } else {
            Err(ParseError::ParserError(format!(
                "Expected {:?}, got {:?}",
                expected, token
            )))
        }
    }

    fn expect_ident(&mut self) -> Result<String, ParseError> {
        match self.peek().clone() {
            Token::Ident(name) => {
                self.pos += 1;
                Ok(name)
            }
            _ => Err(ParseError::ParserError(format!(
                "Expected identifier, got {:?}",
                self.peek()
            ))),
        }
    }

    fn expect_end(&mut self) -> Result<(), ParseError> {
        if self.pos >= self.tokens.len() || self.is_at(Token::Eof) {
            Ok(())
        } else {
            Err(ParseError::ParserError(format!(
                "Unexpected token after expression: {:?}",
                self.peek()
            )))
        }
    }
}

/// 解析表达式（便捷函数）
pub fn parse_expression(tokens: Vec<Token>) -> Result<Expr, ParseError> {
    let mut parser = ExprParser::new(tokens);
    parser.parse()
}
