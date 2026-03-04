//! 函数和过程语句解析 - Function / Sub / Call / Exit

use crate::ast::{Param, Stmt};
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;
use crate::parser::{ParseError, Parser};

impl Parser {
    /// 解析 Function 定义
    pub fn parse_function(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Function)?;
        let name = self.expect_ident()?;

        let params = self.parse_params()?;
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.is_at_end() || self.check_keyword(Keyword::End) {
                break;
            }
            if let Some(stmt) = self.parse_stmt()? {
                body.push(stmt);
            }
            self.skip_newlines();
        }

        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::Function)?;

        Ok(Some(Stmt::Function { name, params, body }))
    }

    /// 解析 Sub 定义
    pub fn parse_sub(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Sub)?;
        let name = self.expect_ident()?;

        let params = self.parse_params()?;
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.is_at_end() || self.check_keyword(Keyword::End) {
                break;
            }
            if let Some(stmt) = self.parse_stmt()? {
                body.push(stmt);
            }
            self.skip_newlines();
        }

        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::Sub)?;

        Ok(Some(Stmt::Sub { name, params, body }))
    }

    /// 解析参数列表
    pub fn parse_params(&mut self) -> Result<Vec<Param>, ParseError> {
        let mut params = vec![];

        if self.match_token(&Token::LParen) {
            if !self.check(&Token::RParen) {
                loop {
                    let is_byref = self.match_keyword(Keyword::ByRef);
                    let _ = self.match_keyword(Keyword::ByVal); // ByVal 是默认的

                    let name = self.expect_ident()?;

                    let default = if self.match_token(&Token::Eq) {
                        Some(self.parse_expr(0)?)
                    } else {
                        None
                    };

                    params.push(Param {
                        name,
                        is_byref,
                        default,
                    });

                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
            }
            self.expect(Token::RParen)?;
        }

        Ok(params)
    }

    /// 解析 Call 语句
    pub fn parse_call(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Call)?;
        let name = self.expect_ident()?;

        let args = if self.match_token(&Token::LParen) {
            let args = self.parse_args()?;
            self.expect(Token::RParen)?;
            args
        } else {
            vec![]
        };

        self.skip_newlines();
        Ok(Some(Stmt::Call { name, args }))
    }

    /// 解析 Exit 语句
    pub fn parse_exit(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Exit)?;

        let stmt = if self.match_keyword(Keyword::For) {
            Stmt::ExitFor
        } else if self.match_keyword(Keyword::Do) {
            Stmt::ExitDo
        } else if self.match_keyword(Keyword::Function) {
            Stmt::ExitFunction
        } else if self.match_keyword(Keyword::Sub) {
            Stmt::ExitSub
        } else if self.match_keyword(Keyword::Property) {
            Stmt::ExitProperty
        } else {
            return Err(ParseError::ParserError(format!(
                "Expected For, Do, Function, Sub, or Property after Exit, got {:?}",
                self.peek()
            )));
        };

        self.skip_newlines();
        Ok(Some(stmt))
    }
}
