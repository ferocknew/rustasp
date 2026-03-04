//! If 语句解析

use crate::ast::IfBranch;
use crate::ast::Stmt;
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析 If 语句
    pub fn parse_if(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::If)?;
        let cond = self.parse_expr(0)?;
        self.skip_newlines();
        self.expect_keyword(Keyword::Then)?;
        self.skip_newlines();

        // 判断是单行 If 还是多行 If
        // 单行 If: if x = 1 then Response.Write("Yes")
        // 多行 If: if x = 1 then \n ... \n end if

        // 检查下一个 token 是否是换行或终止符
        let is_multiline = self.check(&Token::Newline)
            || self.check_keyword(Keyword::End)
            || self.check_keyword(Keyword::Else)
            || self.check_keyword(Keyword::ElseIf);

        if !is_multiline && !self.is_at_end() {
            // 单行 If - 只解析一条语句
            let stmt = self.parse_stmt()?;
            let body = stmt.map_or(vec![], |s| vec![s]);

            // 检查是否有 Else
            let else_block = if self.match_keyword(Keyword::Else) {
                let else_stmt = self.parse_stmt()?;
                Some(else_stmt.map_or(vec![], |s| vec![s]))
            } else {
                None
            };

            return Ok(Some(Stmt::If {
                branches: vec![IfBranch { cond, body }],
                else_block,
            }));
        }

        // 多行 If
        self.skip_newlines();

        let mut branches = vec![IfBranch {
            cond,
            body: vec![],
        }];
        let mut else_block = None;

        // 解析第一个 If 分支的 body
        loop {
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

        // 解析 ElseIf 和 Else 分支
        while !self.check_keyword(Keyword::End) {
            if self.match_keyword(Keyword::ElseIf) {
                let cond = self.parse_expr(0)?;
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
            } else {
                break;
            }
        }

        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::If)?;

        Ok(Some(Stmt::If {
            branches,
            else_block,
        }))
    }
}
