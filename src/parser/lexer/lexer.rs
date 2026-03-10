//! 主 Lexer 结构和 tokenize 逻辑

use super::keyword::lookup_keyword;
use super::token::{SpannedToken, Token};
use crate::parser::ParseError;
use crate::utils::normalize_identifier;

/// 词法分析器
pub struct Lexer {
    pub(crate) input: Vec<char>,
    pub(crate) pos: usize,
    pub(crate) line: usize,
    pub(crate) column: usize,
}

impl Lexer {
    /// 创建新的词法分析器
    pub fn new() -> Self {
        Lexer {
            input: Vec::new(),
            pos: 0,
            line: 1,
            column: 1,
        }
    }

    /// 执行词法分析
    pub fn tokenize(&mut self, source: &str) -> Result<Vec<SpannedToken>, ParseError> {
        self.input = source.chars().collect();
        self.pos = 0;
        self.line = 1;
        self.column = 1;

        let mut tokens = Vec::new();

        while !self.eof() {
            self.skip_whitespace();

            if self.eof() {
                break;
            }

            let token_line = self.line;
            let token_column = self.column;
            let ch = self.current();

            // 注释：从 ' 到行尾
            if ch == '\'' {
                self.skip_comment();
                continue;
            }

            // 换行
            if ch == '\n' {
                self.advance();
                tokens.push(SpannedToken::new(Token::Newline, token_line, token_column));
                continue;
            }

            // 根据字符类型解析 token
            let token = match ch {
                '"' => self.lex_string()?,
                '#' => self.lex_date()?,
                '[' => self.lex_bracket_ident()?,
                c if c.is_ascii_digit() || (c == '.' && self.peek_char(1).is_ascii_digit()) => {
                    self.lex_number()?
                }
                c if Self::is_ident_start(c) => {
                    // 检查是否是 Rem 注释
                    if (c == 'r' || c == 'R') && self.is_rem_keyword() {
                        self.skip_comment();
                        continue;
                    }
                    self.lex_ident_or_keyword()?
                }
                _ => self.lex_operator()?, // 包括 _ 作为续行符在 skip_whitespace 中处理
            };

            tokens.push(SpannedToken::new(token, token_line, token_column));
        }

        tokens.push(SpannedToken::new(Token::Eof, self.line, self.column));
        Ok(tokens)
    }

    /// 检查当前位置是否是 "Rem" 关键字作为注释
    fn is_rem_keyword(&self) -> bool {
        // 检查是否有足够的字符
        if self.pos + 2 >= self.input.len() {
            return false;
        }

        // 检查 "Rem" (不区分大小写)
        let r = self.input[self.pos];
        let e = self.input[self.pos + 1];
        let m = self.input[self.pos + 2];

        (r == 'r' || r == 'R') && (e == 'e' || e == 'E') && (m == 'm' || m == 'M') && {
            // 检查后面是否是非字母字符（确保是完整的单词）
            let next_pos = self.pos + 3;
            next_pos >= self.input.len() || !self.input[next_pos].is_alphanumeric()
        }
    }

    /// 检查字符是否是标识符起始字符
    fn is_ident_start(c: char) -> bool {
        c.is_alphabetic() || c == '_'
    }

    /// 解析字符串字面量
    fn lex_string(&mut self) -> Result<Token, ParseError> {
        self.advance(); // 跳过开始引号
        let mut content = String::new();

        while !self.eof() {
            let ch = self.current();
            if ch == '"' {
                // VBScript 转义: "" 表示一个双引号
                if self.peek_char(1) == '"' {
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

    /// 解析日期字面量
    fn lex_date(&mut self) -> Result<Token, ParseError> {
        self.advance(); // 跳过开始 #
        let content = self.scan_until('#', true)?;
        Ok(Token::Date(content))
    }

    /// 解析数字字面量（包括 .5 格式）
    fn lex_number(&mut self) -> Result<Token, ParseError> {
        let start = self.pos;

        // 处理小数点开头的情况（如 .5）
        if self.current() == '.' {
            self.advance();
            self.scan_while(|c| c.is_ascii_digit());
        } else {
            // 扫描整数部分
            self.scan_while(|c| c.is_ascii_digit());

            // 扫描小数部分
            if self.current() == '.' && self.peek_char(1).is_ascii_digit() {
                self.advance();
                self.scan_while(|c| c.is_ascii_digit());
            }
        }

        let num_str: String = self.input[start..self.pos].iter().collect();
        let num = num_str
            .parse::<f64>()
            .map_err(|_| ParseError::LexerError(format!("Invalid number: {}", num_str)))?;

        Ok(Token::Number(num))
    }

    /// 解析方括号转义的标识符（如 [Error], [Date]）
    fn lex_bracket_ident(&mut self) -> Result<Token, ParseError> {
        self.advance(); // 跳过 [
        let mut content = String::new();

        while !self.eof() {
            let ch = self.current();
            if ch == ']' {
                self.advance();
                // 不检查关键字，方括号内的内容始终作为标识符
                return Ok(Token::Ident(content));
            }
            if ch == '\n' {
                return Err(ParseError::LexerError(
                    "Unterminated bracket identifier".to_string(),
                ));
            }
            content.push(ch);
            self.advance();
        }

        Err(ParseError::LexerError(
            "Unterminated bracket identifier".to_string(),
        ))
    }

    /// 解析标识符或关键字
    fn lex_ident_or_keyword(&mut self) -> Result<Token, ParseError> {
        let start = self.pos;

        self.scan_while(|c| c.is_alphanumeric() || c == '_');

        let ident: String = self.input[start..self.pos].iter().collect();
        let ident_lower = ident.to_ascii_lowercase();

        // 使用关键字查找表
        if let Some(keyword) = lookup_keyword(&ident_lower) {
            Ok(Token::Keyword(keyword))
        } else {
            Ok(Token::Ident(normalize_identifier(&ident)))
        }
    }

    /// 解析运算符和分隔符
    fn lex_operator(&mut self) -> Result<Token, ParseError> {
        let ch = self.current();

        match ch {
            '+' => {
                self.advance();
                Ok(Token::Plus)
            }
            '-' => {
                self.advance();
                Ok(Token::Minus)
            }
            '*' => {
                self.advance();
                Ok(Token::Star)
            }
            '/' => {
                self.advance();
                Ok(Token::Slash)
            }
            '\\' => {
                self.advance();
                Ok(Token::Backslash)
            }
            '^' => {
                self.advance();
                Ok(Token::Caret)
            }
            '&' => {
                self.advance();
                // 检查是否是十六进制或八进制数字
                // VBScript 规则：&H 或 &h 后面必须紧跟十六进制数字
                //              &O 或 &o 后面必须紧跟 0-7 的数字
                //              & 后面紧跟 0-7 且前面是数字或表达式分隔符时才是八进制
                //              否则是字符串连接运算符
                let next = self.current();
                let prev_was_digit_or_expr_end = self.pos > 1
                    && matches!(
                        self.input[self.pos - 2],
                        '0'..='9' | ')' | ']' | '"' | ' ' | '\t' | '(' | ','
                    );

                if (next == 'H' || next == 'h') && self.peek_char(1).is_ascii_hexdigit() {
                    // 十六进制：&HFF&（必须紧跟数字）
                    self.advance(); // 跳过 H
                    let hex_start = self.pos;
                    self.scan_while(|c| c.is_ascii_hexdigit());
                    let hex_str: String = self.input[hex_start..self.pos].iter().collect();

                    // 检查尾随的类型标识符
                    let has_type_suffix = self.current() == '&';
                    if has_type_suffix {
                        self.advance();
                    }

                    let num = u64::from_str_radix(&hex_str, 16).map_err(|_| {
                        ParseError::LexerError(format!("Invalid hexadecimal number: &H{}", hex_str))
                    })?;

                    Ok(Token::Number(num as f64))
                } else if (next == 'O' || next == 'o')
                    && self.peek_char(1).is_ascii_digit()
                    && self.peek_char(1) <= '7'
                {
                    // 八进制：&O77&（必须紧跟 0-7 的数字）
                    self.advance(); // 跳过 O
                    let oct_start = self.pos;
                    self.scan_while(|c| c.is_ascii_digit() && c <= '7');
                    let oct_str: String = self.input[oct_start..self.pos].iter().collect();

                    // 检查尾随的类型标识符
                    let has_type_suffix = self.current() == '&';
                    if has_type_suffix {
                        self.advance();
                    }

                    if oct_str.is_empty() {
                        return Err(ParseError::LexerError(format!(
                            "Invalid octal number at line {}",
                            self.line
                        )));
                    }

                    let num = u64::from_str_radix(&oct_str, 8).map_err(|_| {
                        ParseError::LexerError(format!("Invalid octal number: &O{}", oct_str))
                    })?;

                    Ok(Token::Number(num as f64))
                } else if next.is_ascii_digit() && next <= '7' && prev_was_digit_or_expr_end {
                    // 简化的八进制表示：&77&
                    // 只有在前面是数字或表达式结束符时才解析为八进制
                    let oct_start = self.pos;
                    self.scan_while(|c| c.is_ascii_digit() && c <= '7');
                    let oct_str: String = self.input[oct_start..self.pos].iter().collect();

                    // 检查尾随的类型标识符
                    let has_type_suffix = self.current() == '&';
                    if has_type_suffix {
                        self.advance();
                    }

                    if oct_str.is_empty() {
                        return Err(ParseError::LexerError(format!(
                            "Invalid octal number at line {}",
                            self.line
                        )));
                    }

                    let num = u64::from_str_radix(&oct_str, 8).map_err(|_| {
                        ParseError::LexerError(format!("Invalid octal number: &{}", oct_str))
                    })?;

                    Ok(Token::Number(num as f64))
                } else {
                    // 字符串连接运算符
                    Ok(Token::Ampersand)
                }
            }
            '=' => {
                self.advance();
                Ok(Token::Eq)
            }
            '<' => {
                self.advance();
                let next = self.current();
                if next == '>' {
                    self.advance();
                    Ok(Token::Ne)
                } else if next == '=' {
                    self.advance();
                    Ok(Token::Le)
                } else {
                    Ok(Token::Lt)
                }
            }
            '>' => {
                self.advance();
                if self.current() == '=' {
                    self.advance();
                    Ok(Token::Ge)
                } else {
                    Ok(Token::Gt)
                }
            }
            '(' => {
                self.advance();
                Ok(Token::LParen)
            }
            ')' => {
                self.advance();
                Ok(Token::RParen)
            }
            ',' => {
                self.advance();
                Ok(Token::Comma)
            }
            '.' => {
                self.advance();
                Ok(Token::Dot)
            }
            ':' => {
                self.advance();
                Ok(Token::Colon)
            }
            ';' => {
                // VBScript 不支持分号，跳过
                self.advance();
                Ok(Token::Newline)
            }
            _ => Err(ParseError::LexerError(format!(
                "Unexpected character: '{}' at line {} column {}",
                ch, self.line, self.column
            ))),
        }
    }
}
