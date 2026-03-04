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
            // 检查是否到达语句结束（支持单行 If 语句）
            if self.is_at_end()
                || self.check_keyword(Keyword::End)
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
                    if self.is_at_end()
                        || self.check_keyword(Keyword::End)
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
                    if self.is_at_end() || self.check_keyword(Keyword::End) {
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
            if self.is_at_end() || self.check_keyword(Keyword::Next) {
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
            if self.is_at_end() || self.check_keyword(Keyword::Wend) {
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

    /// 解析 Select Case 语句
    pub(super) fn parse_select(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Select)?;
        self.expect_keyword(Keyword::Case)?;
        
        // 解析 Select Case 的表达式
        let expr = self.parse_expr()?;
        self.skip_newlines();

        let mut cases = vec![];
        let mut else_block = None;

        // 解析所有 Case 分支
        loop {
            // 检查是否到达 End Select
            if self.is_at_end() || self.check_keyword(Keyword::End) {
                break;
            }

            // 期望 Case 关键字
            self.expect_keyword(Keyword::Case)?;
            self.skip_newlines();

            // 检查是否是 Case Else
            if self.check_keyword(Keyword::Else) {
                self.advance();
                self.skip_newlines();
                
                let mut body = vec![];
                loop {
                    if self.is_at_end() || self.check_keyword(Keyword::End) || self.check_keyword(Keyword::Case) {
                        break;
                    }
                    if let Some(stmt) = self.parse_stmt()? {
                        body.push(stmt);
                    }
                    self.skip_newlines();
                }
                else_block = Some(body);
            } else {
                // 解析 Case 的值列表（支持逗号分隔的多个值）
                let mut values = vec![];
                loop {
                    // 解析单个值，遇到逗号或换行符时停止
                    let mut tokens = vec![];
                    while !self.is_at_end() {
                        let next_token = self.peek()?;
                        match next_token {
                            Token::Comma | Token::Newline | Token::Colon => break,
                            Token::Keyword(kw) if matches!(kw, Keyword::End | Keyword::Case) => break,
                            _ => tokens.push(self.advance().clone()),
                        }
                    }
                    
                    if tokens.is_empty() {
                        break;
                    }
                    
                    tokens.push(Token::Eof);
                    let value = crate::parser::expr_parser::parse_expression(tokens)?;
                    values.push(value);
                    
                    // 检查是否有逗号（多个值）
                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
                self.skip_newlines();

                let mut body = vec![];
                loop {
                    if self.is_at_end() || self.check_keyword(Keyword::End) || self.check_keyword(Keyword::Case) {
                        break;
                    }
                    if let Some(stmt) = self.parse_stmt()? {
                        body.push(stmt);
                    }
                    self.skip_newlines();
                }
                
                cases.push(CaseClause {
                    values: Some(values),
                    body,
                });
            }
        }

        // 期望 End Select
        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::Select)?;

        Ok(Some(Stmt::Select {
            expr,
            cases,
            else_block,
        }))
    }
}
