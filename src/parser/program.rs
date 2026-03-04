//! 程序解析 - parse_program

use crate::ast::Stmt;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析程序
    pub fn parse_program(&mut self) -> Result<Vec<Stmt>, ParseError> {
        let mut stmts = vec![];

        while !self.is_at_end() {
            self.skip_newlines();

            if self.is_at_end() {
                break;
            }

            if let Some(stmt) = self.parse_stmt()? {
                stmts.push(stmt);
            }
        }

        Ok(stmts)
    }
}
