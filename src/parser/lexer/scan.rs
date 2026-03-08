//! 扫描工具函数

use crate::parser::ParseError;

impl super::Lexer {
    /// 前进一个字符位置
    pub fn advance(&mut self) {
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

    /// 扫描字符直到满足条件
    pub fn scan_while<F>(&mut self, cond: F)
    where
        F: Fn(char) -> bool,
    {
        while self.pos < self.input.len() && cond(self.current()) {
            self.advance();
        }
    }

    /// 扫描直到遇到结束字符
    pub fn scan_until(&mut self, end: char, error_on_newline: bool) -> Result<String, ParseError> {
        let mut s = String::new();

        while self.pos < self.input.len() {
            let ch = self.current();

            if ch == end {
                self.advance();
                return Ok(s);
            }

            if ch == '\n' && error_on_newline {
                return Err(ParseError::LexerError(format!(
                    "Unterminated literal at line {}",
                    self.line
                )));
            }

            s.push(ch);
            self.advance();
        }

        Err(ParseError::LexerError(format!(
            "Unterminated literal at line {}",
            self.line
        )))
    }

    /// 跳过空白字符
    pub fn skip_whitespace(&mut self) {
        while self.pos < self.input.len() {
            let ch = self.current();

            // 处理续行符: _ 后面跟换行符
            if ch == '_' {
                // 检查后面是否是换行符
                let next_pos = self.pos + 1;
                if next_pos < self.input.len() {
                    let next_ch = self.peek_char(1);
                    if next_ch == '\n' || next_ch == '\r' {
                        // 跳过 _ 和换行符
                        self.advance(); // 跳过 _
                        if next_ch == '\r' && next_pos + 1 < self.input.len()
                            && self.peek_char(2) == '\n' {
                            // Windows 换行 \r\n
                            self.advance(); // 跳过 \r
                            self.advance(); // 跳过 \n
                        } else {
                            self.advance(); // 跳过换行符
                        }
                        // 续行符后可能还有空白，继续跳过
                        continue;
                    }
                }
            }

            if ch == ' ' || ch == '\t' || ch == '\r' {
                self.advance();
            } else {
                break;
            }
        }
    }

    /// 跳过注释（从当前位置到行尾）
    pub fn skip_comment(&mut self) {
        while self.pos < self.input.len() && self.current() != '\n' {
            self.advance();
        }
    }

    /// 检查是否到达文件末尾
    pub fn eof(&self) -> bool {
        self.pos >= self.input.len()
    }

    /// 获取当前字符
    pub fn current(&self) -> char {
        if self.pos < self.input.len() {
            self.input[self.pos]
        } else {
            '\0'
        }
    }

    /// 查看指定偏移位置的字符
    pub fn peek_char(&self, offset: usize) -> char {
        let pos = self.pos + offset;
        if pos < self.input.len() {
            self.input[pos]
        } else {
            '\0'
        }
    }
}
