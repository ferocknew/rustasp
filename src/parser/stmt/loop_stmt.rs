//! 循环语句解析 - For / While / Do

use crate::ast::Expr;
use crate::ast::Stmt;
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析 For 循环
    pub fn parse_for(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::For)?;
        let var = self.expect_ident()?;
        self.expect(Token::Eq)?;
        let start = self.parse_expr(0)?;
        self.expect_keyword(Keyword::To)?;
        let end = self.parse_expr(0)?;

        let step = if self.match_keyword(Keyword::Step) {
            Some(self.parse_expr(0)?)
        } else {
            None
        };
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.is_at_end() || self.check_keyword(Keyword::Next) {
                break;
            }
            match self.parse_stmt()? {
                Some(stmt) => body.push(stmt),
                None => break,
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
    pub fn parse_while(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::While)?;
        let cond = self.parse_expr(0)?;
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.is_at_end() || self.check_keyword(Keyword::Wend) {
                break;
            }
            match self.parse_stmt()? {
                Some(stmt) => body.push(stmt),
                None => break,
            }
            self.skip_newlines();
        }
        self.expect_keyword(Keyword::Wend)?;

        Ok(Some(Stmt::While { cond, body }))
    }

    /// 解析 Do 循环
    pub fn parse_do(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Do)?;
        self.skip_newlines();

        // 检查是 Do While...Loop 还是 Do...Loop While
        let is_while_top = self.match_keyword(Keyword::While);
        let is_until_top = if !is_while_top {
            self.match_keyword(Keyword::Until)
        } else {
            false
        };

        let cond_top: Option<Expr> = if is_while_top || is_until_top {
            // Do While/Until condition ... Loop
            let cond = self.parse_expr(0)?;
            self.skip_newlines();
            Some(cond)
        } else {
            // Do ... Loop While/Until condition
            None
        };

        let mut body = vec![];
        loop {
            if self.is_at_end() || self.check_keyword(Keyword::Loop) {
                break;
            }
            match self.parse_stmt()? {
                Some(stmt) => body.push(stmt),
                None => break,
            }
            self.skip_newlines();
        }

        self.expect_keyword(Keyword::Loop)?;

        // 检查 Loop 后是否有 While/Until
        if cond_top.is_none() {
            let is_while_bottom = self.match_keyword(Keyword::While);
            let is_until_bottom = if !is_while_bottom {
                self.match_keyword(Keyword::Until)
            } else {
                false
            };

            if is_while_bottom || is_until_bottom {
                let cond_expr = self.parse_expr(0)?;
                return Ok(if is_while_bottom {
                    Some(Stmt::DoLoopWhile { body, cond: cond_expr })
                } else {
                    Some(Stmt::DoLoopUntil { body, cond: cond_expr })
                });
            }
        }

        // Do While/Until ... Loop
        let cond = cond_top.unwrap();
        Ok(if is_while_top {
            Some(Stmt::DoWhile { cond, body })
        } else {
            Some(Stmt::DoUntil { cond, body })
        })
    }
}
