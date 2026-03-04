//! 控制流语句解析
//!
//! If, For, While, Select Case

use super::core::StmtParser;
use crate::ast::{CaseClause, IfBranch, Stmt};
use crate::parser::error::ParseError;
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;

impl StmtParser {
    /// 解析 If 语句
    pub(super) fn parse_if(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::If)?;
        let cond = self.parse_expr()?;
        self.skip_newlines();
        self.expect_keyword(Keyword::Then)?;
        self.skip_newlines();

        let mut branches = vec![IfBranch { cond, body: vec![] }];
        let mut else_block = None;

        loop {
            if self.check_keyword(Keyword::End)
                || self.check_keyword(Keyword::Else)
                || self.check_keyword(Keyword::ElseIf)
            {
                break;
            }
            if let Some(stmt) = self.parse_stmt()? {
                branches[0].body.push(stmt);
            }
            self.skip_newlines();
        }

        while !self.check_keyword(Keyword::End) {
            if self.match_keyword(Keyword::ElseIf) {
                let cond = self.parse_expr()?;
                self.skip_newlines();
                self.expect_keyword(Keyword::Then)?;
                self.skip_newlines();

                let mut body = vec![];
                loop {
                    if self.check_keyword(Keyword::End)
                        || self.check_keyword(Keyword::Else)
                        || self.check_keyword(Keyword::ElseIf)
                    {
                        break;
                    }
                    if let Some(stmt) = self.parse_stmt()? {
                        body.push(stmt);
                    }
                    self.skip_newlines();
                }
                branches.push(IfBranch { cond, body });
            } else if self.match_keyword(Keyword::Else) {
                self.skip_newlines();
                let mut body = vec![];
                loop {
                    if self.check_keyword(Keyword::End) {
                        break;
                    }
                    if let Some(stmt) = self.parse_stmt()? {
                        body.push(stmt);
                    }
                    self.skip_newlines();
                }
                else_block = Some(body);
            }
        }

        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::If)?;

        Ok(Some(Stmt::If { branches, else_block }))
    }

    /// 解析 For 循环
    pub(super) fn parse_for(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::For)?;
        let var = self.expect_ident()?;
        self.expect(Token::Eq)?;
        let start = self.parse_expr()?;
        self.expect_keyword(Keyword::To)?;
        let end = self.parse_expr()?;

        let step = if self.match_keyword(Keyword::Step) {
            Some(self.parse_expr()?)
        } else {
            None
        };
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.check_keyword(Keyword::Next) {
                break;
            }
            if let Some(stmt) = self.parse_stmt()? {
                body.push(stmt);
            }
            self.skip_newlines();
        }
        self.expect_keyword(Keyword::Next)?;

        Ok(Some(Stmt::For {
            var,
            start,
            end,
            step,
            body,
        }))
    }

    /// 解析 While 循环
    pub(super) fn parse_while(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::While)?;
        let cond = self.parse_expr()?;
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.check_keyword(Keyword::Wend) {
                break;
            }
            if let Some(stmt) = self.parse_stmt()? {
                body.push(stmt);
            }
            self.skip_newlines();
        }
        self.expect_keyword(Keyword::Wend)?;

        Ok(Some(Stmt::While { cond, body }))
    }
}
