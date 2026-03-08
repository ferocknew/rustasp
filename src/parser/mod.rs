//! Parser 模块 - 统一的语法解析器
//!
//! - Lexer: 手写词法分析器
//! - Parser: 统一解析器（表达式用 Pratt，语句用递归下降）

mod error;

// Lexer 子模块
pub mod lexer;
pub use lexer::token::{Token, SpannedToken};
pub use lexer::keyword::Keyword;

// Parser 核心
mod class_parser;
mod parser;
mod program;

// 表达式解析
mod expr;
mod stmt;

pub use error::ParseError;
pub use lexer::{tokenize, Lexer};
pub use parser::Parser;

/// 解析表达式（便捷函数）
pub fn parse_expr(source: &str) -> Result<crate::ast::Expr, ParseError> {
    let spanned_tokens = tokenize(source)?;
    let mut parser = Parser::with_source(spanned_tokens, source.to_string());
    parser.parse_expr(0)
}

/// 解析程序（便捷函数）
pub fn parse(source: &str) -> Result<crate::ast::Program, ParseError> {
    let spanned_tokens = tokenize(source)?;
    let mut parser = Parser::with_source(spanned_tokens, source.to_string());
    Ok(crate::ast::Program {
        statements: parser.parse_program()?,
    })
}
