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
        self.expect_keyword(Keyword::Then)?;

        // 判断是单行 If 还是多行 If
        let is_multiline = self.check(&Token::Newline);

        if !is_multiline {
            // 单行 If
            let body = self.parse_stmt_list_until(&[Keyword::Else, Keyword::End, Keyword::ElseIf])?;

            let else_block = if self.match_keyword(Keyword::Else) {
                // 跳过 Else 后的换行和冒号（单行 If 的常见写法）
                self.skip_newlines();
                if self.match_token(&Token::Colon) {
                    self.skip_newlines();
                }
                Some(self.parse_stmt_list_until(&[Keyword::End, Keyword::ElseIf])?)
            } else {
                None
            };

            // 可选的 End If
            if self.check_keyword(Keyword::End) {
                self.expect_keyword(Keyword::End)?;
                self.expect_keyword(Keyword::If)?;
            }

            return Ok(Some(Stmt::If {
                branches: vec![IfBranch { cond, body }],
                else_block,
            }));
        }

        // 多行 If
        self.skip_newlines();

        let mut branches = vec![];

        // 解析第一个 If 分支的 body
        let body = self.parse_stmt_list_until(&[Keyword::ElseIf, Keyword::Else, Keyword::End])?;
        branches.push(IfBranch { cond, body });

        // 解析 ElseIf 分支
        while self.match_keyword(Keyword::ElseIf) {
            let cond = self.parse_expr(0)?;
            self.expect_keyword(Keyword::Then)?;
            self.skip_newlines();

            let body = self.parse_stmt_list_until(&[Keyword::ElseIf, Keyword::Else, Keyword::End])?;
            branches.push(IfBranch { cond, body });
        }

        // 解析 Else 分支
        let else_block = if self.match_keyword(Keyword::Else) {
            self.skip_newlines();
            // 跳过 Else 后可能的冒号
            if self.match_token(&Token::Colon) {
                self.skip_newlines();
            }
            Some(self.parse_stmt_list_until(&[Keyword::End])?)
        } else {
            None
        };

        // 结束标记
        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::If)?;

        Ok(Some(Stmt::If { branches, else_block }))
    }
}
