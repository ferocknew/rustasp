//! 后缀表达式解析（成员访问、方法调用、索引访问）

use crate::ast::Expr;
use crate::parser::lexer::Token;
use crate::parser::{ParseError, Parser};

impl Parser {
    /// 解析后缀表达式（成员访问、方法调用、索引访问）
    pub fn parse_postfix(&mut self, mut lhs: Expr) -> Result<Expr, ParseError> {
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
    pub fn parse_args(&mut self) -> Result<Vec<Expr>, ParseError> {
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
}
