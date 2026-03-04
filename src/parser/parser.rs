//! Parser 核心结构
//!
//! 只包含 Parser struct 和基础工具方法

use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;
use crate::parser::ParseError;

/// 解析器
pub struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    /// 创建新的解析器
    pub fn new(tokens: Vec<Token>) -> Self {
        Parser { tokens, pos: 0 }
    }

    // ==================== 核心工具方法 ====================

    /// 查看当前 token
    pub fn peek(&self) -> &Token {
        if self.pos < self.tokens.len() {
            &self.tokens[self.pos]
        } else {
            &Token::Eof
        }
    }

    /// 查看后面的 token（用于 lookahead）
    pub fn peek_ahead(&self, offset: usize) -> &Token {
        let pos = self.pos + offset;
        if pos < self.tokens.len() {
            &self.tokens[pos]
        } else {
            &Token::Eof
        }
    }

    /// 前进一个位置并返回 token
    pub fn advance(&mut self) -> &Token {
        if self.pos < self.tokens.len() {
            let token = &self.tokens[self.pos];
            self.pos += 1;
            token
        } else {
            &Token::Eof
        }
    }

    /// 是否到达末尾
    pub fn is_at_end(&self) -> bool {
        self.pos >= self.tokens.len() || matches!(self.tokens[self.pos], Token::Eof)
    }

    /// 检查当前 token 是否匹配
    pub fn check(&self, token: &Token) -> bool {
        std::mem::discriminant(self.peek()) == std::mem::discriminant(token)
    }

    /// 检查是否是某个关键字
    pub fn check_keyword(&self, keyword: Keyword) -> bool {
        matches!(self.peek(), Token::Keyword(k) if *k == keyword)
    }

    /// 尝试匹配 token
    pub fn match_token(&mut self, token: &Token) -> bool {
        if self.check(token) {
            self.advance();
            true
        } else {
            false
        }
    }

    /// 尝试匹配关键字
    pub fn match_keyword(&mut self, keyword: Keyword) -> bool {
        if self.check_keyword(keyword) {
            self.advance();
            true
        } else {
            false
        }
    }

    /// 期望某个 token
    pub fn expect(&mut self, token: Token) -> Result<&Token, ParseError> {
        if self.check(&token) {
            Ok(self.advance())
        } else {
            Err(ParseError::ParserError(format!(
                "Expected {:?}, got {:?}",
                token,
                self.peek()
            )))
        }
    }

    /// 期望某个关键字
    pub fn expect_keyword(&mut self, keyword: Keyword) -> Result<(), ParseError> {
        if self.check_keyword(keyword) {
            self.advance();
            Ok(())
        } else {
            Err(ParseError::ParserError(format!(
                "Expected keyword {:?}, got {:?}",
                keyword,
                self.peek()
            )))
        }
    }

    /// 期望标识符
    pub fn expect_ident(&mut self) -> Result<String, ParseError> {
        match self.peek().clone() {
            Token::Ident(name) => {
                self.advance();
                Ok(name)
            }
            _ => Err(ParseError::ParserError(format!(
                "Expected identifier, got {:?}",
                self.peek()
            ))),
        }
    }

    /// 跳过换行符
    pub fn skip_newlines(&mut self) {
        while self.match_token(&Token::Newline) {}
    }

    /// 获取当前位置
    pub fn pos(&self) -> usize {
        self.pos
    }
}
