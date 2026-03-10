//! 前缀表达式解析

use crate::ast::{Expr, UnaryOp};
use crate::parser::Keyword;
use crate::parser::lexer::Token;
use crate::parser::{ParseError, Parser};

impl Parser {
    /// 解析前缀表达式
    pub fn parse_prefix(&mut self) -> Result<Expr, ParseError> {
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
            Token::Date(date_str) => {
                self.advance();
                Ok(Expr::Date(date_str))
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

            // True 和 False 关键字作为布尔值
            Token::Keyword(Keyword::True) => {
                self.advance();
                Ok(Expr::Boolean(true))
            }
            Token::Keyword(Keyword::False) => {
                self.advance();
                Ok(Expr::Boolean(false))
            }

            // Nothing 关键字
            Token::Keyword(Keyword::Nothing) => {
                self.advance();
                Ok(Expr::Null)
            }
            // Null 关键字
            Token::Keyword(Keyword::Null) => {
                self.advance();
                Ok(Expr::Null)
            }
            // Empty 关键字
            Token::Keyword(Keyword::Empty) => {
                self.advance();
                Ok(Expr::Empty)
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

            // With 上下文中的成员访问：.property 或 .method(...)
            Token::Dot => {
                self.advance(); // 消耗 .
                // 点号后面允许标识符或关键字作为属性名
                let name = match self.peek().clone() {
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

                // 检查是否是方法调用（有括号）
                if self.check(&Token::LParen) {
                    self.advance(); // 消耗 (
                    let args = self.parse_args_in_parens()?;
                    self.expect(Token::RParen)?;
                    Ok(Expr::WithMethod {
                        method: name,
                        args,
                    })
                } else {
                    // 先创建 WithProperty，后续可能在 postfix 中转换为 WithMethod
                    self.parse_postfix(Expr::WithProperty { property: name })
                }
            }

            _ => Err(self.error_with_context(format!(
                "Unexpected token in expression: {:?}",
                token
            ))),
        }
    }

    /// 一元运算符的绑定优先级
    fn unary_binding_power(&self) -> u8 {
        17 // 高于所有二元运算符
    }
}
