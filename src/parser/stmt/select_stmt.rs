//! Select Case 语句解析

use crate::ast::CaseClause;
use crate::ast::Stmt;
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析 Select Case 语句
    pub fn parse_select(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Select)?;
        self.expect_keyword(Keyword::Case)?;

        let expr = self.parse_expr(0)?;
        self.skip_newlines();

        let mut cases = vec![];
        let mut else_block = None;

        loop {
            if self.is_at_end() || self.check_keyword(Keyword::End) {
                break;
            }

            self.expect_keyword(Keyword::Case)?;

            if self.match_keyword(Keyword::Else) {
                self.skip_newlines();
                let mut body = vec![];
                loop {
                    if self.is_at_end()
                        || self.check_keyword(Keyword::End)
                        || self.check_keyword(Keyword::Case)
                    {
                        break;
                    }
                    match self.parse_stmt()? {
                        Some(stmt) => {
                            body.push(stmt);

                            // 处理冒号分隔的后续语句
                            while self.match_token(&Token::Colon) {
                                self.skip_newlines();

                                if self.is_at_end()
                                    || self.check(&Token::Newline)
                                    || self.check_keyword(Keyword::End)
                                    || self.check_keyword(Keyword::Case)
                                {
                                    break;
                                }

                                if let Some(next_stmt) = self.parse_stmt()? {
                                    body.push(next_stmt);
                                }
                            }
                        }
                        None => break,
                    }
                    self.skip_newlines();
                }
                else_block = Some(body);
            } else {
                // 解析 Case 值列表
                let mut values = vec![];
                loop {
                    let value = self.parse_expr(0)?;
                    values.push(value);

                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
                self.skip_newlines();

                let mut body = vec![];
                loop {
                    if self.is_at_end()
                        || self.check_keyword(Keyword::End)
                        || self.check_keyword(Keyword::Case)
                    {
                        break;
                    }
                    match self.parse_stmt()? {
                        Some(stmt) => {
                            body.push(stmt);

                            // 处理冒号分隔的后续语句
                            while self.match_token(&Token::Colon) {
                                self.skip_newlines();

                                if self.is_at_end()
                                    || self.check(&Token::Newline)
                                    || self.check_keyword(Keyword::End)
                                    || self.check_keyword(Keyword::Case)
                                {
                                    break;
                                }

                                if let Some(next_stmt) = self.parse_stmt()? {
                                    body.push(next_stmt);
                                }
                            }
                        }
                        None => break,
                    }
                    self.skip_newlines();
                }

                cases.push(CaseClause {
                    values: Some(values),
                    body,
                });
            }
        }

        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::Select)?;

        Ok(Some(Stmt::Select {
            expr,
            cases,
            else_block,
        }))
    }
}
