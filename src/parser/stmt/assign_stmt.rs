//! 赋值语句和表达式语句解析

use crate::ast::Stmt;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析标识符开头的语句（赋值或表达式）
    pub fn parse_ident_stmt(&mut self) -> Result<Option<Stmt>, ParseError> {
        // 解析左侧表达式
        let lhs = self.parse_expr(0)?;

        // 检查是否是赋值
        if self.check(&crate::parser::lexer::Token::Eq) {
            self.advance(); // 消耗 =
            let value = self.parse_expr(0)?;
            self.skip_newlines();
            return Ok(Some(Stmt::Assignment {
                target: lhs,
                value,
            }));
        }

        // 否则是表达式语句
        self.skip_newlines();
        Ok(Some(Stmt::Expr(lhs)))
    }

    /// 解析表达式语句
    pub fn parse_expr_stmt(&mut self) -> Result<Option<Stmt>, ParseError> {
        let expr = self.parse_expr(0)?;
        self.skip_newlines();
        Ok(Some(Stmt::Expr(expr)))
    }

    /// 解析 Set 语句
    pub fn parse_set(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(crate::parser::keyword::Keyword::Set)?;
        let target = self.parse_expr(0)?;
        self.expect(crate::parser::lexer::Token::Eq)?;
        let value = self.parse_expr(0)?;
        self.skip_newlines();

        Ok(Some(Stmt::Set { target, value }))
    }
}
