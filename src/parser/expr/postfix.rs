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
                    // 点号后面允许标识符或关键字作为属性名（如 Response.End）
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
                    lhs = Expr::Property {
                        object: Box::new(lhs),
                        property: name,
                    };
                }

                // 方法调用或索引访问：(args)
                Token::LParen => {
                    self.advance(); // 消耗 (
                    let args = self.parse_args_in_parens()?;
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
                            // 其他情况作为索引处理（支持多维索引）
                            Expr::Index {
                                object: Box::new(lhs),
                                indices: args,
                            }
                        }
                    };
                }

                // 不再是后缀操作符，结束循环
                _ => break,
            }
        }

        // 循环结束后，检查是否是 VBScript 风格的无括号调用
        // 支持：MyFunc "Hello" 或 Response.Write "Hello"
        if self.is_arg_start() {
            let args = self.parse_args_no_parens()?;
            lhs = match lhs {
                Expr::Property { object, property } => Expr::Method {
                    object,
                    method: property,
                    args,
                },
                Expr::Variable(name) => Expr::Call { name, args },
                _ => lhs,
            };
        }

        Ok(lhs)
    }

    /// 解析括号内的参数列表（用逗号分隔）
    /// 调用者已消耗 LParen，需要返回前消耗 RParen
    fn parse_args_in_parens(&mut self) -> Result<Vec<Expr>, ParseError> {
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

    /// 解析无括号的参数列表（用空格或逗号分隔）
    /// 用于 VBScript 风格的调用：MyFunc arg1, arg2
    fn parse_args_no_parens(&mut self) -> Result<Vec<Expr>, ParseError> {
        let mut args = vec![];

        while self.is_arg_start() {
            args.push(self.parse_expr(0)?);
            // 检查是否有逗号或更多参数
            if !self.match_token(&Token::Comma) {
                // 没有逗号，检查是否还有更多参数（空格分隔）
                if !self.is_arg_start() {
                    break;
                }
            }
        }

        Ok(args)
    }

    /// 解析参数列表（保留用于兼容性）
    ///
    /// 支持两种模式：
    /// 1. 括号内：func(arg1, arg2) - 用逗号分隔
    /// 2. 无括号：func arg1 arg2 - 用空格分隔
    pub fn parse_args(&mut self) -> Result<Vec<Expr>, ParseError> {
        // 检查是否在括号内（用于决定分隔符）
        if self.check(&Token::LParen) {
            self.advance(); // 消耗 (
            let args = self.parse_args_in_parens()?;
            self.expect(Token::RParen)?;
            Ok(args)
        } else {
            self.parse_args_no_parens()
        }
    }

    /// 检查是否是参数的开始
    fn is_arg_start(&self) -> bool {
        match self.peek() {
            Token::String(_) => true,
            Token::Number(_) => true,
            Token::Ident(_) => true,
            _ => false,
        }
    }
}
