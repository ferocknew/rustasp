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

            let start_pos = self.pos();
            match self.parse_stmt()? {
                Some(stmt) => stmts.push(stmt),
                None => {
                    // 如果没有解析到语句且位置没有前进，报错避免死循环
                    if self.pos() == start_pos && !self.is_at_end() {
                        return Err(ParseError::ParserError(format!(
                            "Unexpected token at position {}: {:?}",
                            self.pos(),
                            self.peek()
                        )));
                    }
                }
            }
        }

        Ok(stmts)
    }
}
