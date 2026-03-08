//! 解析错误

use thiserror::Error;

#[derive(Debug, Error)]
pub enum ParseError {
    #[error("Lexer error: {0}")]
    LexerError(String),

    #[error("Parser error: {0}")]
    ParserError(String),

    #[error("Unexpected token: expected {expected}, found {found}")]
    UnexpectedToken { expected: String, found: String },

    #[error("Unexpected end of input")]
    UnexpectedEnd,

    #[error("Invalid syntax: {0}")]
    InvalidSyntax(String),

    /// 带上下文的解析错误
    #[error("Parser error at position {pos}: {message}\n\nToken context:\n{context}")]
    ParserErrorWithContext {
        pos: usize,
        message: String,
        context: String,
    },
}

impl ParseError {
    /// 创建带上下文的错误
    pub fn with_context(message: String, pos: usize, context: String) -> Self {
        ParseError::ParserErrorWithContext {
            pos,
            message,
            context,
        }
    }
}
