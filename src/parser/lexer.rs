//! 词法分析器（手写实现，不依赖 chumsky）

use super::error::ParseError;
use super::keyword::Keyword;
use serde::{Deserialize, Serialize};

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

/// 词法分析器
pub struct Lexer {
    input: Vec<char>,
    pos: usize,
    line: usize,
    column: usize,
}

impl Lexer {
    /// 创建词法分析器
    pub fn new() -> Self {
        Lexer {
            input: Vec::new(),
            pos: 0,
            line: 1,
            column: 1,
        }
    }

    /// 解析源代码为 Token 列表
    pub fn tokenize(&mut self, source: &str) -> Result<Vec<Token>, ParseError> {
        self.input = source.chars().collect();
        self.pos = 0;
        self.line = 1;
        self.column = 1;

        let mut tokens = Vec::new();

        while self.pos < self.input.len() {
            self.skip_whitespace();

            if self.pos >= self.input.len() {
                break;
            }

            let ch = self.input[self.pos];

            // 注释：从 ' 到行尾
            if ch == '\'' {
                self.skip_comment();
                continue;
            }

            // 换行
            if ch == '\n' {
                tokens.push(Token::Newline);
                self.advance();
                self.line += 1;
                self.column = 1;
                continue;
            }

            // 字符串
            if ch == '"' {
                tokens.push(self.lex_string()?);
                continue;
            }

            // 数字
            if ch.is_ascii_digit() {
                tokens.push(self.lex_number()?);
                continue;
            }

            // 标识符或关键字
            if ch.is_alphabetic() || ch == '_' {
                tokens.push(self.lex_ident_or_keyword()?);
                continue;
            }

            // 运算符和分隔符
            tokens.push(self.lex_operator_or_delim()?);
        }

        tokens.push(Token::Eof);
        Ok(tokens)
    }

    fn advance(&mut self) {
        if self.pos < self.input.len() {
            if self.input[self.pos] == '\n' {
                self.line += 1;
                self.column = 1;
            } else {
                self.column += 1;
            }
            self.pos += 1;
        }
    }

    fn current(&self) -> char {
        if self.pos < self.input.len() {
            self.input[self.pos]
        } else {
            '\0'
        }
    }

    fn peek(&self, offset: usize) -> char {
        let pos = self.pos + offset;
        if pos < self.input.len() {
            self.input[pos]
        } else {
            '\0'
        }
    }

    fn skip_whitespace(&mut self) {
        while self.pos < self.input.len() {
            let ch = self.input[self.pos];
            if ch == ' ' || ch == '\t' || ch == '\r' {
                self.advance();
            } else {
                break;
            }
        }
    }

    fn skip_comment(&mut self) {
        while self.pos < self.input.len() && self.input[self.pos] != '\n' {
            self.advance();
        }
    }

    fn lex_string(&mut self) -> Result<Token, ParseError> {
        self.advance(); // 跳过开始引号
        let mut content = String::new();

        while self.pos < self.input.len() {
            let ch = self.input[self.pos];
            if ch == '"' {
                // VBScript 转义: "" 表示一个双引号
                if self.peek(1) == '"' {
                    // 两个双引号 -> 一个双引号
                    content.push('"');
                    self.advance(); // 跳过第一个 "
                    self.advance(); // 跳过第二个 "
                } else {
                    // 单个双引号 -> 字符串结束
                    self.advance();
                    return Ok(Token::String(content));
                }
            } else if ch == '\n' {
                return Err(ParseError::LexerError(format!(
                    "Unterminated string at line {}",
                    self.line
                )));
            } else {
                content.push(ch);
                self.advance();
            }
        }

        Err(ParseError::LexerError(format!(
            "Unterminated string at line {}",
            self.line
        )))
    }

    fn lex_number(&mut self) -> Result<Token, ParseError> {
        let start = self.pos;
        let mut has_dot = false;

        while self.pos < self.input.len() {
            let ch = self.input[self.pos];
            if ch.is_ascii_digit() {
                self.advance();
            } else if ch == '.' && !has_dot && self.peek(1).is_ascii_digit() {
                has_dot = true;
                self.advance();
            } else {
                break;
            }
        }

        let num_str: String = self.input[start..self.pos].iter().collect();
        let num = num_str
            .parse::<f64>()
            .map_err(|_| ParseError::LexerError(format!("Invalid number: {}", num_str)))?;

        Ok(Token::Number(num))
    }

    fn lex_ident_or_keyword(&mut self) -> Result<Token, ParseError> {
        let start = self.pos;

        while self.pos < self.input.len() {
            let ch = self.input[self.pos];
            if ch.is_alphanumeric() || ch == '_' {
                self.advance();
            } else {
                break;
            }
        }

        let ident: String = self.input[start..self.pos].iter().collect();
        let ident_lower = ident.to_lowercase();

        // 特殊处理布尔值
        match ident_lower.as_str() {
            "true" => return Ok(Token::Boolean(true)),
            "false" => return Ok(Token::Boolean(false)),
            _ => {}
        }

        // 检查是否是关键字
        if let Some(keyword) = self.match_keyword(&ident) {
            Ok(Token::Keyword(keyword))
        } else {
            Ok(Token::Ident(ident))
        }
    }

    fn match_keyword(&self, ident: &str) -> Option<Keyword> {
        let ident_lower = ident.to_lowercase();
        match ident_lower.as_str() {
            "dim" => Some(Keyword::Dim),
            "const" => Some(Keyword::Const),
            "if" => Some(Keyword::If),
            "then" => Some(Keyword::Then),
            "else" => Some(Keyword::Else),
            "elseif" => Some(Keyword::ElseIf),
            "end" => Some(Keyword::End),
            "for" => Some(Keyword::For),
            "to" => Some(Keyword::To),
            "step" => Some(Keyword::Step),
            "next" => Some(Keyword::Next),
            "each" => Some(Keyword::Each),
            "in" => Some(Keyword::In),
            "do" => Some(Keyword::Do),
            "while" => Some(Keyword::While),
            "loop" => Some(Keyword::Loop),
            "until" => Some(Keyword::Until),
            "wend" => Some(Keyword::Wend),
            "exit" => Some(Keyword::Exit),
            "sub" => Some(Keyword::Sub),
            "function" => Some(Keyword::Function),
            "call" => Some(Keyword::Call),
            "set" => Some(Keyword::Set),
            "let" => Some(Keyword::Let),
            "class" => Some(Keyword::Class),
            "property" => Some(Keyword::Property),
            "get" => Some(Keyword::Get),
            "public" => Some(Keyword::Public),
            "private" => Some(Keyword::Private),
            "and" => Some(Keyword::And),
            "or" => Some(Keyword::Or),
            "not" => Some(Keyword::Not),
            "xor" => Some(Keyword::Xor),
            "mod" => Some(Keyword::Mod),
            "is" => Some(Keyword::Is),
            "true" | "false" => None, // 布尔值由调用方处理
            _ => None,
        }
    }

    fn lex_operator_or_delim(&mut self) -> Result<Token, ParseError> {
        let ch = self.current();

        let token = match ch {
            '+' => {
                self.advance();
                Token::Plus
            }
            '-' => {
                self.advance();
                Token::Minus
            }
            '*' => {
                self.advance();
                Token::Star
            }
            '/' => {
                self.advance();
                Token::Slash
            }
            '\\' => {
                self.advance();
                Token::Backslash
            }
            '^' => {
                self.advance();
                Token::Caret
            }
            '&' => {
                self.advance();
                Token::Ampersand
            }
            '<' => {
                if self.peek(1) == '>' {
                    self.advance();
                    self.advance();
                    Token::Ne
                } else if self.peek(1) == '=' {
                    self.advance();
                    self.advance();
                    Token::Le
                } else {
                    self.advance();
                    Token::Lt
                }
            }
            '>' => {
                if self.peek(1) == '=' {
                    self.advance();
                    self.advance();
                    Token::Ge
                } else {
                    self.advance();
                    Token::Gt
                }
            }
            '=' => {
                self.advance();
                Token::Eq
            }
            '(' => {
                self.advance();
                Token::LParen
            }
            ')' => {
                self.advance();
                Token::RParen
            }
            ',' => {
                self.advance();
                Token::Comma
            }
            '.' => {
                self.advance();
                Token::Dot
            }
            ':' => {
                self.advance();
                Token::Colon
            }
            _ => {
                return Err(ParseError::LexerError(format!(
                    "Unexpected character '{}' at line {}, column {}",
                    ch, self.line, self.column
                )))
            }
        };

        Ok(token)
    }
}

impl Default for Lexer {
    fn default() -> Self {
        Self::new()
    }
}

/// 解析源代码为 Token 列表（便捷函数）
pub fn tokenize(source: &str) -> Result<Vec<Token>, ParseError> {
    let mut lexer = Lexer::new();
    lexer.tokenize(source)
}
