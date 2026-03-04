//! 声明语句解析
//!
//! Dim, Const, Set, Call, Exit, Option

use super::core::StmtParser;
use crate::ast::Stmt;
use crate::parser::error::ParseError;
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;

impl StmtParser {
    /// 解析 Dim 语句
    pub(super) fn parse_dim(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Dim)?;
        let names = self.parse_dim_var_list()?;

        if names.len() == 1 {
            self.skip_newlines();
            if self.match_token(&Token::Eq) {
                self.skip_newlines();
                let init = Some(self.parse_expr()?);
                return Ok(Some(Stmt::Dim {
                    name: names.into_iter().next().unwrap(),
                    init,
                    is_array: false,
                    sizes: vec![],
                }));
            }
        }

        let mut names_iter = names.into_iter();
        let first = names_iter.next().unwrap();

        for name in names_iter {
            self.pending_dims.push(Stmt::Dim {
                name,
                init: None,
                is_array: false,
                sizes: vec![],
            });
        }

        Ok(Some(Stmt::Dim {
            name: first,
            init: None,
            is_array: false,
            sizes: vec![],
        }))
    }

    fn parse_dim_var_list(&mut self) -> Result<Vec<String>, ParseError> {
        let mut names = vec![self.expect_ident()?];
        while self.match_token(&Token::Comma) {
            names.push(self.expect_ident()?);
        }
        Ok(names)
    }

    /// 解析 Const 语句
    pub(super) fn parse_const(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Const)?;
        let name = self.expect_ident()?;
        self.expect(Token::Eq)?;
        let value = self.parse_expr()?;
        Ok(Some(Stmt::Const { name, value }))
    }

    /// 解析 Option 语句（如 Option Explicit）
    pub(super) fn parse_option(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Option)?;
        
        if self.match_keyword(Keyword::Explicit) {
            Ok(Some(Stmt::OptionExplicit))
        } else {
            Err(ParseError::ParserError(
                "Expected Explicit after Option".into(),
            ))
        }
    }

    /// 解析 Set 语句
    pub(super) fn parse_set(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Set)?;
        let target = self.parse_expr()?;
        self.expect(Token::Eq)?;
        let value = self.parse_expr()?;
        Ok(Some(Stmt::Set { target, value }))
    }

    /// 解析 Call 语句
    pub(super) fn parse_call(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Call)?;
        let name = self.expect_ident()?;

        let mut args = vec![];
        if self.match_token(&Token::LParen) {
            if !self.check(&Token::RParen) {
                loop {
                    args.push(self.parse_expr()?);
                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
            }
            self.expect(Token::RParen)?;
        }

        Ok(Some(Stmt::Call { name, args }))
    }

    /// 解析 Exit 语句
    pub(super) fn parse_exit(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Exit)?;

        if self.match_keyword(Keyword::For) {
            Ok(Some(Stmt::ExitFor))
        } else if self.match_keyword(Keyword::Do) {
            Ok(Some(Stmt::ExitDo))
        } else if self.match_keyword(Keyword::Function) {
            Ok(Some(Stmt::ExitFunction))
        } else if self.match_keyword(Keyword::Sub) {
            Ok(Some(Stmt::ExitSub))
        } else {
            Err(ParseError::ParserError(
                "Expected For, Do, Function, or Sub after Exit".into(),
            ))
        }
    }
}
