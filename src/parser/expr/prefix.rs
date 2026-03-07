//! 前缀表达式解析

use crate::ast::{Expr, UnaryOp};
use crate::parser::keyword::Keyword;
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

    /// 一元运算符的绑定优先级
    fn unary_binding_power(&self) -> u8 {
        17 // 高于所有二元运算符
    }
}
