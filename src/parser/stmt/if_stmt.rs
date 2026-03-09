//! If 语句解析

use crate::ast::IfBranch;
use crate::ast::Stmt;
use crate::parser::Keyword;
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
        // 单行 If (冒号分隔): if x = 1 then : Response.Write("Yes") : End If
        // 单行 If (带 Else): if x = 1 then : a = 1 : Else : a = 2 : End If
        // 多行 If: if x = 1 then \n ... \n end if
        let is_multiline = self.check(&Token::Newline);

        self.skip_newlines();  // 跳过 Then 后的换行

        if !is_multiline && !self.is_at_end() && !self.check_keyword(Keyword::End) {
            // 单行 If - 支持冒号分隔的多条语句
            // 例如: If x > 0 Then Response.Write "yes" : Response.Write "<br>"
            // 或者: If x > 0 Then : Response.Write "yes" : End If
            // 或者: If x > 0 Then : Response.Write "yes" : Else : Response.Write "no" : End If
            let mut body = vec![];

            // 如果 Then 后面是冒号，先消耗它
            if self.check(&Token::Colon) {
                // 检查下一个 token，如果是 End，说明这是空的 If 块
                // 例如: If False Then : End If
                self.advance(); // 消耗冒号
                self.skip_newlines();

                // 检查是否是空 If 块
                if self.check_keyword(Keyword::End) {
                    // 空 If 块: If False Then : End If
                    self.expect_keyword(Keyword::End)?;
                    self.expect_keyword(Keyword::If)?;
                    return Ok(Some(Stmt::If {
                        branches: vec![IfBranch { cond, body: vec![] }],
                        else_block: None,
                    }));
                }
            }

            // 解析第一条语句
            if let Some(stmt) = self.parse_stmt()? {
                body.push(stmt);

                // 处理冒号分隔的后续语句（VBScript 语法糖）
                loop {
                    self.skip_newlines();

                    // 检查是否到达 End If
                    if self.check_keyword(Keyword::End) {
                        if self.peek_next_is_keyword(Keyword::If) {
                            // 这是带 End If 的单行 If
                            break;
                        }
                    }

                    // 检查是否遇到 Else
                    if self.check_keyword(Keyword::Else) {
                        // 单行 If 的 Else 分支
                        self.expect_keyword(Keyword::Else)?;
                        self.skip_newlines();

                        let mut else_body = vec![];

                        // 检查 Else 后是否有冒号
                        if self.check(&Token::Colon) {
                            self.advance(); // 消耗冒号
                            self.skip_newlines();
                        }

                        // 解析 Else 分支的语句
                        loop {
                            // 检查是否到达 End If
                            if self.check_keyword(Keyword::End) && self.peek_next_is_keyword(Keyword::If) {
                                break;
                            }

                            if let Some(stmt) = self.parse_stmt()? {
                                else_body.push(stmt);
                            }

                            self.skip_newlines();

                            // 检查是否有冒号分隔符
                            if !self.match_token(&Token::Colon) {
                                break;
                            }

                            self.skip_newlines();
                        }

                        // 消耗 End If
                        self.expect_keyword(Keyword::End)?;
                        self.expect_keyword(Keyword::If)?;

                        return Ok(Some(Stmt::If {
                            branches: vec![IfBranch { cond, body }],
                            else_block: Some(else_body),
                        }));
                    }

                    // 检查是否遇到冒号
                    if !self.match_token(&Token::Colon) {
                        // 没有冒号了，检查其他结束条件
                        // 注意：遇到 Else 时不要 break，让第69行处理
                        if self.is_at_end()
                            || self.check(&Token::Newline)
                            || self.check_keyword(Keyword::ElseIf)
                            || self.check_keyword(Keyword::End)
                        {
                            break;
                        }
                        // 如果遇到 Else，继续循环让第69行处理
                        if self.check_keyword(Keyword::Else) {
                            continue;
                        }
                        continue;
                    }

                    // 跳过冒号后的空白
                    self.skip_newlines();

                    // 检查是否到达行尾或文件结束
                    // 注意：不检查 Else，因为第69行会处理 Else
                    if self.is_at_end()
                        || self.check(&Token::Newline)
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
            } else {
                // 空语句块: If False Then : End If
                // 检查是否有 End If
                if self.check_keyword(Keyword::End) && self.peek_next_is_keyword(Keyword::If) {
                    self.expect_keyword(Keyword::End)?;
                    self.expect_keyword(Keyword::If)?;
                    return Ok(Some(Stmt::If {
                        branches: vec![IfBranch { cond, body: vec![] }],
                        else_block: None,
                    }));
                }
            }

            // 检查是否有 End If (单行 If 也可能带 End If)
            if self.check_keyword(Keyword::End) && self.peek_next_is_keyword(Keyword::If) {
                self.expect_keyword(Keyword::End)?;
                self.expect_keyword(Keyword::If)?;
            }

            // 单行 If 语句不支持 Else 块
            // 在 VBScript 中，如果要使用 Else，必须使用多行 If 语句
            // 如果在这里遇到 Else，说明这个 Else 是属于外层的，不应该被消费
            return Ok(Some(Stmt::If {
                branches: vec![IfBranch { cond, body }],
                else_block: None,
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
                Some(stmt) => {
                    branches[0].body.push(stmt);

                    // 处理冒号分隔的后续语句（VBScript 语法糖）
                    while self.match_token(&Token::Colon) {
                        self.skip_newlines();

                        if self.is_at_end()
                            || self.check(&Token::Newline)
                            || self.check_keyword(Keyword::End)
                            || self.check_keyword(Keyword::Else)
                            || self.check_keyword(Keyword::ElseIf)
                        {
                            break;
                        }

                        if let Some(next_stmt) = self.parse_stmt()? {
                            branches[0].body.push(next_stmt);
                        }
                    }
                }
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
                        Some(stmt) => {
                            body.push(stmt);

                            // 处理冒号分隔的后续语句
                            while self.match_token(&Token::Colon) {
                                self.skip_newlines();

                                if self.is_at_end()
                                    || self.check(&Token::Newline)
                                    || self.check_keyword(Keyword::End)
                                    || self.check_keyword(Keyword::Else)
                                    || self.check_keyword(Keyword::ElseIf)
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
                        Some(stmt) => {
                            body.push(stmt);

                            // 处理冒号分隔的后续语句
                            while self.match_token(&Token::Colon) {
                                self.skip_newlines();

                                if self.is_at_end()
                                    || self.check(&Token::Newline)
                                    || self.check_keyword(Keyword::End)
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
