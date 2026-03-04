//! Parser 模块 - 词法分析 + 表达式解析
//!
//! - Lexer: 手写词法分析器
//! - ExprParser: Pratt 算法表达式解析器

mod error;
pub mod expr_parser;
pub mod keyword;
pub mod lexer;

pub use error::ParseError;
pub use expr_parser::{parse_expression, ExprParser};
pub use lexer::{tokenize, Lexer, Token};

/// 解析表达式（便捷函数）
pub fn parse_expr(source: &str) -> Result<crate::ast::Expr, ParseError> {
    let tokens = tokenize(source)?;
    parse_expression(tokens)
}
