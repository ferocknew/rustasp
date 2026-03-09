//! Select Case 语句解析

use crate::ast::CaseClause;
use crate::ast::Stmt;
use crate::parser::Keyword;
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
            if self.is_at_end() {
                break;
            }

            // 检查是否是 End Select（外层 Select 的结束）
            if self.check_keyword(Keyword::End) {
                // 向前看一个 token，检查是否是 End Select
                let pos = self.pos();
                self.advance(); // 消耗 End
                if self.check_keyword(Keyword::Select) {
                    // 这是 End Select，回退并退出循环
                    self.seek_to(pos);
                    break;
                } else {
                    // 不是 End Select，回退
                    self.seek_to(pos);
                }
            }

            self.expect_keyword(Keyword::Case)?;

            if self.match_keyword(Keyword::Else) {
                self.skip_newlines();

                // 解析 Case Else 的 body
                // 注意：需要正确处理嵌套的 Select Case 和 If...End If
                let body = self.parse_case_body()?;

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

                // 解析 Case 的 body
                let body = self.parse_case_body()?;

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

    /// 解析 Case body（正确处理嵌套结构）
    fn parse_case_body(&mut self) -> Result<Vec<Stmt>, ParseError> {
        let mut body = vec![];

        loop {
            self.skip_newlines();

            // 检查是否结束
            if self.is_at_end() {
                break;
            }

            // 检查是否遇到 End Select（需要向前看）
            if self.check_keyword(Keyword::End) {
                let pos = self.pos();
                self.advance(); // 消耗 End

                if self.check_keyword(Keyword::Select) {
                    // 这是 End Select，回退并退出
                    self.seek_to(pos);
                    break;
                } else if self.check_keyword(Keyword::If)
                    || self.check_keyword(Keyword::Function)
                    || self.check_keyword(Keyword::Sub)
                    || self.check_keyword(Keyword::Class)
                {
                    // 这是其他 End 语句（End If, End Function, End Sub, End Class）
                    // 回退，让 parse_stmt 正常处理
                    self.seek_to(pos);
                } else {
                    // 未知的 End 组合，回退并退出（安全起见）
                    self.seek_to(pos);
                    break;
                }
            }

            // 检查是否遇到下一个 Case（但不是 Case Else）
            if self.check_keyword(Keyword::Case) {
                let pos = self.pos();
                self.advance(); // 消耗 Case
                if self.check_keyword(Keyword::Else) {
                    // 这是 Case Else，回退并退出
                    self.seek_to(pos);
                    break;
                } else {
                    // 这是普通的 Case，回退并退出
                    self.seek_to(pos);
                    break;
                }
            }

            // 解析一条语句
            if let Some(stmt) = self.parse_stmt()? {
                body.push(stmt);
            }
        }

        Ok(body)
    }
}
