//! VBScript Token 定义

use serde::{Deserialize, Serialize};

/// VBScript Token 类型
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Token {
    // 字面量
    String(String),
    Number(f64),
    Boolean(bool),

    // 标识符和关键字
    Ident(String),
    Keyword(Keyword),

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
    Comma,
    Dot,
    Colon,

    // 特殊
    Newline,
    Eof,
}

/// VBScript 关键字
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Keyword {
    Dim,
    Const,
    If,
    Then,
    Else,
    ElseIf,
    End,
    For,
    To,
    Step,
    Next,
    Do,
    While,
    Loop,
    Until,
    Exit,
    Sub,
    Function,
    Call,
    Return,
    Set,
    Let,
    Class,
    Property,
    Get,
    Let_,
    Public,
    Private,
    True,
    False,
    Nothing,
    Empty,
    Null,
    And,
    Or,
    Not,
    Xor,
    Mod,
    Is,
    In,
    Option,
    Explicit,
    On,
    Error,
    Resume,
    Next_,
    ReDim,
    Preserve,
    Erase,
    Execute,
    ExecuteGlobal,
    Eval,
}
