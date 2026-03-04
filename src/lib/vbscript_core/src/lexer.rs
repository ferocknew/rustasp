//! VBScript 词法分析器

use chumsky::prelude::*;
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

/// 词法分析器
pub type Lexer = impl Parser<char, Vec<Token>, Error = Simple<char>>;

/// 创建词法分析器
pub fn lexer() -> Lexer {
    let comment = just('\'').then_ignore(take_until(just('\n').or(end())));

    let string = just('"')
        .ignore_then(filter(|c| *c != '"').repeated())
        .then_ignore(just('"'))
        .collect::<String>()
        .map(Token::String);

    let number = filter(|c: &char| c.is_ascii_digit())
        .repeated()
        .chain(just('.'))
        .chain(filter(|c: &char| c.is_ascii_digit()).repeated())
        .or(filter(|c: &char| c.is_ascii_digit()).repeated())
        .collect::<String>()
        .from_str()
        .unwrapped()
        .map(Token::Number);

    let boolean = text::keyword("true")
        .map(|_| Token::Boolean(true))
        .or(text::keyword("false").map(|_| Token::Boolean(false)));

    let keyword = text::keyword("dim").map(|_| Token::Keyword(Keyword::Dim))
        .or(text::keyword("const").map(|_| Token::Keyword(Keyword::Const)))
        .or(text::keyword("if").map(|_| Token::Keyword(Keyword::If)))
        .or(text::keyword("then").map(|_| Token::Keyword(Keyword::Then)))
        .or(text::keyword("else").map(|_| Token::Keyword(Keyword::Else)))
        .or(text::keyword("elseif").map(|_| Token::Keyword(Keyword::ElseIf)))
        .or(text::keyword("end").map(|_| Token::Keyword(Keyword::End)))
        .or(text::keyword("for").map(|_| Token::Keyword(Keyword::For)))
        .or(text::keyword("to").map(|_| Token::Keyword(Keyword::To)))
        .or(text::keyword("step").map(|_| Token::Keyword(Keyword::Step)))
        .or(text::keyword("next").map(|_| Token::Keyword(Keyword::Next)))
        .or(text::keyword("do").map(|_| Token::Keyword(Keyword::Do)))
        .or(text::keyword("while").map(|_| Token::Keyword(Keyword::While)))
        .or(text::keyword("loop").map(|_| Token::Keyword(Keyword::Loop)))
        .or(text::keyword("until").map(|_| Token::Keyword(Keyword::Until)))
        .or(text::keyword("sub").map(|_| Token::Keyword(Keyword::Sub)))
        .or(text::keyword("function").map(|_| Token::Keyword(Keyword::Function)))
        .or(text::keyword("call").map(|_| Token::Keyword(Keyword::Call)))
        .or(text::keyword("set").map(|_| Token::Keyword(Keyword::Set)))
        .or(text::keyword("class").map(|_| Token::Keyword(Keyword::Class)))
        .or(text::keyword("public").map(|_| Token::Keyword(Keyword::Public)))
        .or(text::keyword("private").map(|_| Token::Keyword(Keyword::Private)))
        .or(text::keyword("and").map(|_| Token::Keyword(Keyword::And)))
        .or(text::keyword("or").map(|_| Token::Keyword(Keyword::Or)))
        .or(text::keyword("not").map(|_| Token::Keyword(Keyword::Not)))
        .or(text::keyword("mod").map(|_| Token::Keyword(Keyword::Mod)));

    let ident = text::ident().map(|s: String| Token::Ident(s));

    let op = just('+').map(|_| Token::Plus)
        .or(just('-').map(|_| Token::Minus))
        .or(just('*').map(|_| Token::Star))
        .or(just('/').map(|_| Token::Slash))
        .or(just('\\').map(|_| Token::Backslash))
        .or(just('^').map(|_| Token::Caret))
        .or(just('&').map(|_| Token::Ampersand))
        .or(just::<_, _, Simple<char>>("<>").map(|_| Token::Ne))
        .or(just("<=").map(|_| Token::Le))
        .or(just(">=").map(|_| Token::Ge))
        .or(just('<').map(|_| Token::Lt))
        .or(just('>').map(|_| Token::Gt))
        .or(just('=').map(|_| Token::Eq));

    let delim = just('(').map(|_| Token::LParen)
        .or(just(')').map(|_| Token::RParen))
        .or(just(',').map(|_| Token::Comma))
        .or(just('.').map(|_| Token::Dot))
        .or(just(':').map(|_| Token::Colon));

    let newline = just('\n').map(|_| Token::Newline);

    let token = comment
        .ignore_then(newline.clone())
        .or(comment.ignore_then(end()))
        .or(string)
        .or(number)
        .or(boolean)
        .or(keyword)
        .or(ident)
        .or(op)
        .or(delim)
        .or(newline);

    token
        .recover_with(skip_then_retry_until([]))
        .padded_by(just(' ').or(just('\t')).repeated())
        .repeated()
        .then_ignore(end())
}

/// 解析源代码为 Token 列表
pub fn tokenize(source: &str) -> Result<Vec<Token>, String> {
    lexer()
        .parse(source.chars().collect::<Vec<_>>())
        .map_err(|e| format!("Lexer error: {:?}", e))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_tokenize_string() {
        let tokens = tokenize(r#""Hello, World!""#).unwrap();
        assert_eq!(tokens, vec![Token::String("Hello, World!".to_string())]);
    }

    #[test]
    fn test_tokenize_number() {
        let tokens = tokenize("42").unwrap();
        assert_eq!(tokens, vec![Token::Number(42.0)]);
    }

    #[test]
    fn test_tokenize_ident() {
        let tokens = tokenize("Response").unwrap();
        assert_eq!(tokens, vec![Token::Ident("Response".to_string())]);
    }
}
