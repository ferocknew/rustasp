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

        // VBScript 的 If 语句统一处理：不区分单行/多行
        self.skip_newlines();

        let mut branches = vec![];

        // 解析第一个 If 分支的 body
        let body = self.parse_stmt_list_until_if()?;
        branches.push(IfBranch { cond, body });

        // 解析 ElseIf 分支
        while self.match_keyword(Keyword::ElseIf) {
            let cond = self.parse_expr(0)?;
            self.expect_keyword(Keyword::Then)?;
            self.skip_newlines();

            let body = self.parse_stmt_list_until_if()?;
            branches.push(IfBranch { cond, body });
        }

        // 解析 Else 分支
        let else_block = if self.match_keyword(Keyword::Else) {
            self.skip_newlines();
            Some(self.parse_stmt_list_until_else()?)
        } else {
            None
        };

        // 结束标记（可选，单行 If 可以省略）
        if self.check_keyword(Keyword::End) && matches!(self.peek_ahead(1), Token::Keyword(Keyword::If)) {
            self.expect_keyword(Keyword::End)?;
            self.expect_keyword(Keyword::If)?;
        }

        Ok(Some(Stmt::If { branches, else_block }))
    }
}
