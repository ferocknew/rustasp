//! 解析器辅助方法
//!
//! Token 消费、查找等辅助方法

use super::core::StmtParser;
use crate::parser::error::ParseError;
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;

impl StmtParser {
    pub(super) fn peek(&self) -> Result<&Token, ParseError> {
        Ok(if self.pos < self.tokens.len() {
            &self.tokens[self.pos]
        } else {
            &Token::Eof
        })
    }

    pub(super) fn peek_ahead(&self, offset: usize) -> Result<&Token, ParseError> {
        let pos = self.pos + offset;
        Ok(if pos < self.tokens.len() {
            &self.tokens[pos]
        } else {
            &Token::Eof
        })
    }

    pub(super) fn advance(&mut self) -> &Token {
        let token = if self.pos < self.tokens.len() {
            &self.tokens[self.pos]
        } else {
            &Token::Eof
        };
        self.pos += 1;
        token
    }

    pub(super) fn is_at_end(&self) -> bool {
        self.pos >= self.tokens.len() || matches!(self.tokens[self.pos], Token::Eof)
    }

    pub(super) fn check(&self, token: &Token) -> bool {
        std::mem::discriminant(self.peek().unwrap()) == std::mem::discriminant(token)
    }

    pub(super) fn check_keyword(&self, keyword: Keyword) -> bool {
        matches!(self.peek().unwrap(), Token::Keyword(k) if *k == keyword)
    }

    pub(super) fn match_token(&mut self, token: &Token) -> bool {
        if self.check(token) {
            self.advance();
            true
        } else {
            false
        }
    }

    pub(super) fn match_keyword(&mut self, keyword: Keyword) -> bool {
        if self.check_keyword(keyword) {
            self.advance();
            true
        } else {
            false
        }
    }

    pub(super) fn expect(&mut self, token: Token) -> Result<&Token, ParseError> {
        if self.check(&token) {
            Ok(self.advance())
        } else {
            Err(ParseError::ParserError(format!(
                "Expected {:?}, got {:?}",
                token,
                self.peek()?
            )))
        }
    }

    pub(super) fn expect_keyword(&mut self, keyword: Keyword) -> Result<(), ParseError> {
        if self.check_keyword(keyword) {
            self.advance();
            Ok(())
        } else {
            Err(ParseError::ParserError(format!(
                "Expected keyword {:?}, got {:?}",
                keyword,
                self.peek()?
            )))
        }
    }

    pub(super) fn expect_ident(&mut self) -> Result<String, ParseError> {
        match self.peek()?.clone() {
            Token::Ident(name) => {
                self.advance();
                Ok(name)
            }
            _ => Err(ParseError::ParserError(format!(
                "Expected identifier, got {:?}",
                self.peek()?
            ))),
        }
    }

    pub(super) fn skip_newlines(&mut self) {
        while self.match_token(&Token::Newline) {}
    }
}
