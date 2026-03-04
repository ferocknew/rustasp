//! 语句解析器模块
//!
//! 递归下降解析器，将 Token 流解析为 AST 语句。

mod control;
mod core;
mod declarations;
mod helpers;
mod procedures;

pub use core::{parse_program, StmtParser};
