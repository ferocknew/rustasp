//! Parser 模块 - 统一的语法解析器
//!
//! - Lexer: 手写词法分析器
//! - Parser: 统一解析器（表达式用 Pratt，语句用递归下降）

mod error;
pub mod keyword;
pub mod lexer;

// 子模块
mod parser;
mod expr;
mod stmt;

pub use error::ParseError;
pub use keyword::Keyword;
pub use lexer::{tokenize, Lexer, Token};
pub use parser::Parser;

/// 解析表达式（便捷函数）
pub fn parse_expr(source: &str) -> Result<crate::ast::Expr, ParseError> {
    let tokens = tokenize(source)?;
    let mut parser = Parser::new(tokens);
    parser.parse_expr()
}

/// 解析程序（便捷函数）
pub fn parse(source: &str) -> Result<crate::ast::Program, ParseError> {
    let tokens = tokenize(source)?;
    let mut parser = Parser::new(tokens);
    parser.parse_program()
}
