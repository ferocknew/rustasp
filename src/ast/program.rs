//! 程序定义

use serde::{Deserialize, Serialize};
use super::Stmt;

/// 程序（顶层 AST）
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Program {
    pub statements: Vec<Stmt>,
}

impl Program {
    pub fn new() -> Self {
        Program { statements: Vec::new() }
    }

    pub fn with_statements(statements: Vec<Stmt>) -> Self {
        Program { statements }
    }

    pub fn push(&mut self, stmt: Stmt) {
        self.statements.push(stmt);
    }
}

impl Default for Program {
    fn default() -> Self {
        Self::new()
    }
}
