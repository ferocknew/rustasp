//! Select Case 语句解析

use crate::ast::CaseClause;
use crate::ast::Stmt;
use crate::parser::lexer::Token;
use crate::parser::Keyword;
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

        // 循环解析每个 Case 分支
        loop {
            self.skip_newlines();

            // 检查是否到达 End Select
            if self.check_keyword(Keyword::End) {
                break;
            }

            if self.is_at_end() {
                return Err(ParseError::ParserError(
                    "Unexpected end of file in Select Case".to_string(),
                ));
            }

            // 期望 Case 关键字
            self.expect_keyword(Keyword::Case)?;

            // 检查是否是 Case Else
            if self.match_keyword(Keyword::Else) {
                self.skip_newlines();
                // Case Else 的 body 解析到 End 为止
                let body = self.parse_stmt_list_until(&[Keyword::End])?;
                else_block = Some(body);
                break;
            }

            // 解析 Case 值列表（支持多个值，用逗号分隔）
            let mut values = vec![];
            loop {
                let value = self.parse_expr(0)?;
                values.push(value);

                if !self.match_token(&Token::Comma) {
                    break;
                }
            }
            self.skip_newlines();

            // Case body 解析到下一个 Case 或 End 为止
            let body = self.parse_stmt_list_until(&[Keyword::Case, Keyword::End])?;

            cases.push(CaseClause {
                values: Some(values),
                body,
            });
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
