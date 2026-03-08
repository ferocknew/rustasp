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

        // 判断是单行 If 还是多行 If
        // 单行 If: if x = 1 then Response.Write("Yes")
        // 多行 If: if x = 1 then \n ... \n end if
        let is_multiline = self.check(&Token::Newline);

        self.skip_newlines();  // 跳过 Then 后的换行

        if !is_multiline && !self.is_at_end() && !self.check_keyword(Keyword::End) {
            // 单行 If - 支持冒号分隔的多条语句
            // 例如: If x > 0 Then Response.Write "yes" : Response.Write "<br>"
            let mut body = vec![];

            // 解析第一条语句
            if let Some(stmt) = self.parse_stmt()? {
                body.push(stmt);

                // 处理冒号分隔的后续语句（VBScript 语法糖）
                while self.match_token(&Token::Colon) {
                    // 跳过冒号后的空白
                    self.skip_newlines();

                    // 检查是否到达行尾或文件结束
                    if self.is_at_end()
                        || self.check(&Token::Newline)
                        || self.check_keyword(Keyword::Else)
                        || self.check_keyword(Keyword::End)
                        || self.check_keyword(Keyword::ElseIf)
                    {
                        break;
                    }

                    // 解析下一条语句
                    if let Some(next_stmt) = self.parse_stmt()? {
                        body.push(next_stmt);
                    }
                }
            }

            // 检查是否有 Else
            let else_block = if self.match_keyword(Keyword::Else) {
                let mut else_body = vec![];

                // 解析 Else 后的第一条语句
                if let Some(stmt) = self.parse_stmt()? {
                    else_body.push(stmt);

                    // 处理冒号分隔的后续语句
                    while self.match_token(&Token::Colon) {
                        self.skip_newlines();

                        if self.is_at_end() || self.check(&Token::Newline) {
                            break;
                        }

                        if let Some(next_stmt) = self.parse_stmt()? {
                            else_body.push(next_stmt);
                        }
                    }
                }

                if else_body.is_empty() {
                    None
                } else {
                    Some(else_body)
                }
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
            self.skip_newlines();  // 先跳过换行

            if self.is_at_end()
                || self.check_keyword(Keyword::End)
                || self.check_keyword(Keyword::Else)
                || self.check_keyword(Keyword::ElseIf)
            {
                break;
            }

            match self.parse_stmt()? {
                Some(stmt) => branches[0].body.push(stmt),
                None => break,  // 如果没有解析到语句，退出循环
            }
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
                    self.skip_newlines();

                    if self.is_at_end()
                        || self.check_keyword(Keyword::End)
                        || self.check_keyword(Keyword::Else)
                        || self.check_keyword(Keyword::ElseIf)
                    {
                        break;
                    }
                    match self.parse_stmt()? {
                        Some(stmt) => body.push(stmt),
                        None => break,
                    }
                }
                branches.push(IfBranch { cond, body });
            } else if self.match_keyword(Keyword::Else) {
                self.skip_newlines();
                let mut body = vec![];
                loop {
                    self.skip_newlines();

                    if self.is_at_end() || self.check_keyword(Keyword::End) {
                        break;
                    }
                    match self.parse_stmt()? {
                        Some(stmt) => body.push(stmt),
                        None => break,
                    }
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
