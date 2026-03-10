//! Token 定义

use serde::{Deserialize, Serialize};

/// VBScript Token
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Token {
    // 字面量
    String(String),
    Number(f64),
    Boolean(bool),
    Date(String), // 日期字面量，如 #2024-01-01#

    // 标识符和关键字
    Ident(String),
    Keyword(super::keyword::Keyword),

    // 运算符
    Plus,
    Minus,
    Star,
    Slash,
    Backslash,
    Caret,
    Ampersand,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,

    // 分隔符
    LParen,
    RParen,
    LeftBracket,  // [ 用于转义关键字
    RightBracket, // ] 用于转义关键字
    Comma,
    Dot,
    Colon,

    // 特殊
    Newline,
    Eof,

    // 特殊值（可选，也可以用关键字表示）
    Null,
    Empty,
}

/// 带位置信息的 Token
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SpannedToken {
    pub token: Token,
    pub line: usize,
    pub column: usize,
}

impl SpannedToken {
    pub fn new(token: Token, line: usize, column: usize) -> Self {
        Self {
            token,
            line,
            column,
        }
    }
}
