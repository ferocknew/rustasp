//! Parser 模块 - 语法解析层
//!
//! 输入源码字符串，输出 AST，不执行代码

mod keyword;
mod lexer;
mod parser;
mod error;

pub use keyword::Keyword;
pub use lexer::{Token, Lexer};
pub use parser::Parser;
pub use error::ParseError;

/// 解析源代码为 AST
pub fn parse(source: &str) -> Result<crate::ast::Program, ParseError> {
    let tokens = lexer::tokenize(source)?;
    parser::parse_tokens(&tokens)
}

/// 解析表达式
pub fn parse_expr(source: &str) -> Result<crate::ast::Expr, ParseError> {
    let tokens = lexer::tokenize(source)?;
    parser::parse_expr_tokens(&tokens)
}
