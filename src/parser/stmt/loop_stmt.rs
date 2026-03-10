//! 循环语句解析 - For / For Each / While / Do

use crate::ast::Expr;
use crate::ast::Stmt;
use crate::parser::lexer::Token;
use crate::parser::Keyword;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析 For 循环（可能是 For 或 For Each）
    pub fn parse_for(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::For)?;

        // 检查是否是 For Each
        if self.match_keyword(Keyword::Each) {
            return self.parse_for_each();
        }

        // 普通 For 循环
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

        let body = self.parse_block_until(&[Keyword::Next])?;

        self.expect_keyword(Keyword::Next)?;

        Ok(Some(Stmt::For {
            var,
            start,
            end,
            step,
            body,
        }))
    }

    /// 解析 For Each 循环
    fn parse_for_each(&mut self) -> Result<Option<Stmt>, ParseError> {
        let var = self.expect_ident()?;
        self.expect_keyword(Keyword::In)?;
        let collection = self.parse_expr(0)?;
        self.skip_newlines();

        let body = self.parse_block_until(&[Keyword::Next])?;

        self.expect_keyword(Keyword::Next)?;

        Ok(Some(Stmt::ForEach {
            var,
            collection,
            body,
        }))
    }

    /// 解析 While 循环
    pub fn parse_while(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::While)?;
        let cond = self.parse_expr(0)?;
        self.skip_newlines();

        let body = self.parse_block_until(&[Keyword::Wend])?;

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

        let body = self.parse_block_until(&[Keyword::Loop])?;

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
                    Some(Stmt::DoLoopWhile {
                        body,
                        cond: cond_expr,
                    })
                } else {
                    Some(Stmt::DoLoopUntil {
                        body,
                        cond: cond_expr,
                    })
                });
            } else {
                // Do ... Loop (无限循环)
                return Ok(Some(Stmt::DoLoop { body }));
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
