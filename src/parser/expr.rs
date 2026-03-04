//! 表达式解析 - Pratt 算法
//!
//! 使用绑定优先级（binding power）实现，直接操作 Parser 的 pos

use crate::ast::{BinaryOp, Expr, UnaryOp};
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;
use crate::parser::{ParseError, Parser};

impl Parser {
    /// 解析表达式（入口函数）
    pub fn parse_expr(&mut self, min_bp: u8) -> Result<Expr, ParseError> {
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

            // 消耗运算符
            let op_token = self.advance().clone();
            let op = self.token_to_binary_op(&op_token)?;

            // 解析右侧表达式（使用右侧优先级）
            let rhs = self.parse_expr(r_bp)?;

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
        let token = self.peek().clone();

        match token {
            // 字面量
            Token::Number(n) => {
                self.advance();
                Ok(Expr::Number(n))
            }
            Token::String(s) => {
                self.advance();
                Ok(Expr::String(s))
            }
            Token::Boolean(b) => {
                self.advance();
                Ok(Expr::Boolean(b))
            }
            Token::Null => {
                self.advance();
                Ok(Expr::Null)
            }
            Token::Empty => {
                self.advance();
                Ok(Expr::Empty)
            }

            // 变量或标识符
            Token::Ident(name) => {
                self.advance();
                self.parse_postfix(Expr::Variable(name))
            }

            // 一元运算符
            Token::Minus => {
                self.advance();
                let rhs = self.parse_expr(self.unary_binding_power())?;
                Ok(Expr::Unary {
                    op: UnaryOp::Neg,
                    operand: Box::new(rhs),
                })
            }

            Token::Keyword(kw) if kw.is_unary_op() => {
                self.advance();
                let rhs = self.parse_expr(self.unary_binding_power())?;
                Ok(Expr::Unary {
                    op: UnaryOp::Not,
                    operand: Box::new(rhs),
                })
            }

            // New 关键字
            Token::Keyword(Keyword::New) => {
                self.advance();
                let name = self.expect_ident()?;
                Ok(Expr::New(name))
            }

            // 括号表达式
            Token::LParen => {
                self.advance();
                let expr = self.parse_expr(0)?;
                self.expect(Token::RParen)?;
                // 括号表达式后也可能有后缀
                self.parse_postfix(expr)
            }

            _ => Err(ParseError::ParserError(format!(
                "Unexpected token in expression: {:?}",
                token
            ))),
        }
    }

    /// 解析后缀表达式（成员访问、方法调用、索引访问）
    fn parse_postfix(&mut self, mut lhs: Expr) -> Result<Expr, ParseError> {
        loop {
            match self.peek() {
                // 成员访问：obj.property 或 obj.method(...)
                Token::Dot => {
                    self.advance(); // 消耗 .
                    let name = self.expect_ident()?;
                    lhs = Expr::Property {
                        object: Box::new(lhs),
                        property: name,
                    };
                }

                // 方法调用或索引访问：(args)
                Token::LParen => {
                    self.advance(); // 消耗 (
                    let args = self.parse_args()?;
                    self.expect(Token::RParen)?;

                    // 如果 lhs 是 Property，转换为 Method
                    lhs = match lhs {
                        Expr::Property { object, property } => Expr::Method {
                            object,
                            method: property,
                            args,
                        },
                        Expr::Variable(name) => Expr::Call { name, args },
                        _ => {
                            // 其他情况作为索引处理
                            if args.len() == 1 {
                                Expr::Index {
                                    object: Box::new(lhs),
                                    index: Box::new(args.into_iter().next().unwrap()),
                                }
                            } else {
                                return Err(ParseError::ParserError(
                                    "Invalid index expression".to_string(),
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

    /// 解析参数列表
    fn parse_args(&mut self) -> Result<Vec<Expr>, ParseError> {
        let mut args = vec![];

        if !self.check(&Token::RParen) {
            loop {
                args.push(self.parse_expr(0)?);
                if !self.match_token(&Token::Comma) {
                    break;
                }
            }
        }

        Ok(args)
    }

    /// 获取中缀运算符的绑定优先级
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
            Token::Caret => Some((16, 15)),

            // Mod
            Token::Keyword(Keyword::Mod) => Some((13, 14)),

            // Is 运算符
            Token::Keyword(Keyword::Is) => Some((5, 6)),

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
            Token::Keyword(Keyword::Mod) => Ok(BinaryOp::Mod),
            Token::Keyword(Keyword::Is) => Ok(BinaryOp::Is),
            _ => Err(ParseError::ParserError(format!(
                "Expected binary operator, got {:?}",
                token
            ))),
        }
    }
}
