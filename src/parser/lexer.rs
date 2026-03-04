//! 词法分析器

use chumsky::prelude::*;
use serde::{Deserialize, Serialize};
use super::keyword::Keyword;
use super::error::ParseError;

/// VBScript Token
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
    Plus, Minus, Star, Slash, Backslash, Caret, Ampersand,
    Eq, Ne, Lt, Le, Gt, Ge,

    // 分隔符
    LParen, RParen, Comma, Dot, Colon,

    // 特殊
    Newline,
    Eof,
}

/// 词法分析器
pub struct Lexer;

impl Lexer {
    /// 创建词法分析器
    pub fn new() -> Self {
        Lexer
    }

    /// 解析源代码为 Token 列表
    pub fn tokenize(&self, source: &str) -> Result<Vec<Token>, ParseError> {
        tokenize(source)
    }
}

impl Default for Lexer {
    fn default() -> Self {
        Self::new()
    }
}

/// 解析源代码为 Token 列表
pub fn tokenize(source: &str) -> Result<Vec<Token>, ParseError> {
    let lexer = create_lexer();
    lexer
        .parse(source.chars().collect::<Vec<_>>())
        .map_err(|e| ParseError::LexerError(format!("{:?}", e)))
}

fn create_lexer() -> impl Parser<char, Vec<Token>, Error = Simple<char>> {
    let comment = just('\'').then_ignore(take_until(just('\n').or(end())));

    let string = just('"')
        .ignore_then(filter(|c| *c != '"' && *c != '\n').repeated())
        .then_ignore(just('"'))
        .collect::<String>()
        .map(Token::String);

    let number = filter(|c: &char| c.is_ascii_digit())
        .repeated()
        .chain(just('.').chain(filter(|c: &char| c.is_ascii_digit()).repeated()).or_not().flatten())
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
        .or(text::keyword("each").map(|_| Token::Keyword(Keyword::Each)))
        .or(text::keyword("in").map(|_| Token::Keyword(Keyword::In)))
        .or(text::keyword("do").map(|_| Token::Keyword(Keyword::Do)))
        .or(text::keyword("while").map(|_| Token::Keyword(Keyword::While)))
        .or(text::keyword("loop").map(|_| Token::Keyword(Keyword::Loop)))
        .or(text::keyword("until").map(|_| Token::Keyword(Keyword::Until)))
        .or(text::keyword("wend").map(|_| Token::Keyword(Keyword::Wend)))
        .or(text::keyword("exit").map(|_| Token::Keyword(Keyword::Exit)))
        .or(text::keyword("sub").map(|_| Token::Keyword(Keyword::Sub)))
        .or(text::keyword("function").map(|_| Token::Keyword(Keyword::Function)))
        .or(text::keyword("call").map(|_| Token::Keyword(Keyword::Call)))
        .or(text::keyword("set").map(|_| Token::Keyword(Keyword::Set)))
        .or(text::keyword("let").map(|_| Token::Keyword(Keyword::Let)))
        .or(text::keyword("class").map(|_| Token::Keyword(Keyword::Class)))
        .or(text::keyword("property").map(|_| Token::Keyword(Keyword::Property)))
        .or(text::keyword("get").map(|_| Token::Keyword(Keyword::Get)))
        .or(text::keyword("public").map(|_| Token::Keyword(Keyword::Public)))
        .or(text::keyword("private").map(|_| Token::Keyword(Keyword::Private)))
        .or(text::keyword("and").map(|_| Token::Keyword(Keyword::And)))
        .or(text::keyword("or").map(|_| Token::Keyword(Keyword::Or)))
        .or(text::keyword("not").map(|_| Token::Keyword(Keyword::Not)))
        .or(text::keyword("xor").map(|_| Token::Keyword(Keyword::Xor)))
        .or(text::keyword("mod").map(|_| Token::Keyword(Keyword::Mod)))
        .or(text::keyword("is").map(|_| Token::Keyword(Keyword::Is)))
        .or(text::keyword("option").map(|_| Token::Keyword(Keyword::Option)))
        .or(text::keyword("explicit").map(|_| Token::Keyword(Keyword::Explicit)))
        .or(text::keyword("on").map(|_| Token::Keyword(Keyword::On)))
        .or(text::keyword("error").map(|_| Token::Keyword(Keyword::Error)))
        .or(text::keyword("resume").map(|_| Token::Keyword(Keyword::Resume)))
        .or(text::keyword("redim").map(|_| Token::Keyword(Keyword::ReDim)))
        .or(text::keyword("preserve").map(|_| Token::Keyword(Keyword::Preserve)))
        .or(text::keyword("erase").map(|_| Token::Keyword(Keyword::Erase)))
        .or(text::keyword("execute").map(|_| Token::Keyword(Keyword::Execute)))
        .or(text::keyword("executeglobal").map(|_| Token::Keyword(Keyword::ExecuteGlobal)))
        .or(text::keyword("eval").map(|_| Token::Keyword(Keyword::Eval)))
        .or(text::keyword("select").map(|_| Token::Keyword(Keyword::Select)))
        .or(text::keyword("case").map(|_| Token::Keyword(Keyword::Case)));

    let ident = text::ident().map(|s: String| Token::Ident(s));

    let op = just('+').map(|_| Token::Plus)
        .or(just('-').map(|_| Token::Minus))
        .or(just('*').map(|_| Token::Star))
        .or(just('/').map(|_| Token::Slash))
        .or(just('\\').map(|_| Token::Backslash))
        .or(just('^').map(|_| Token::Caret))
        .or(just('&').map(|_| Token::Ampersand))
        .or(just("<>").map(|_| Token::Ne))
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

    let token = comment.ignore_then(newline.clone())
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
        .padded_by(just(' ').or(just('\t')).or(just('\r')).repeated())
        .repeated()
        .then_ignore(end())
}
