//! 过程定义语句解析
//!
//! Function, Sub

use super::core::StmtParser;
use crate::ast::{Param, Stmt};
use crate::parser::error::ParseError;
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;

impl StmtParser {
    /// 解析 Function 定义
    pub(super) fn parse_function(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Function)?;
        let name = self.expect_ident()?;

        let mut params = vec![];
        if self.match_token(&Token::LParen) {
            if !self.check(&Token::RParen) {
                loop {
                    params.push(Param::new(self.expect_ident()?));
                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
            }
            self.expect(Token::RParen)?;
        }
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.check_keyword(Keyword::End) && matches!(self.peek_ahead(1)?, Token::Keyword(Keyword::Function)) {
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
    pub(super) fn parse_sub(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Sub)?;
        let name = self.expect_ident()?;

        let mut params = vec![];
        if self.match_token(&Token::LParen) {
            if !self.check(&Token::RParen) {
                loop {
                    params.push(Param::new(self.expect_ident()?));
                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
            }
            self.expect(Token::RParen)?;
        }
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.check_keyword(Keyword::End) && matches!(self.peek_ahead(1)?, Token::Keyword(Keyword::Sub)) {
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
}
