//! 声明语句解析 - Dim / Const / Option / ReDim

use crate::ast::Stmt;
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析 Dim 声明（支持逗号分隔的多变量）
    pub fn parse_dim(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Dim)?;

        // 解析第一个变量
        let name = self.expect_ident()?;

        // 检查是否是数组
        let is_array = self.match_token(&Token::LParen);
        let mut sizes = vec![];

        if is_array {
            if !self.check(&Token::RParen) {
                loop {
                    sizes.push(self.parse_expr(0)?);
                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
            }
            self.expect(Token::RParen)?;
        }

        // 检查初始化
        let init = if self.check(&Token::Eq) && !self.check(&Token::Colon) {
            self.advance();
            Some(self.parse_expr(0)?)
        } else {
            None
        };

        // 检查是否有逗号分隔的更多变量
        if self.match_token(&Token::Comma) {
            // 有更多变量，跳过它们
            loop {
                self.skip_newlines();
                match self.peek() {
                    Token::Ident(_) => {
                        // 跳过变量名
                        self.advance();
                        // 跳过数组声明（如果有）
                        if self.match_token(&Token::LParen) {
                            while !self.check(&Token::RParen) && !self.is_at_end() {
                                self.advance();
                            }
                            self.expect(Token::RParen)?;
                        }
                        // 跳过初始化（如果有）
                        if self.match_token(&Token::Eq) {
                            // 跳过初始化表达式
                            while !self.check(&Token::Comma) && !self.check(&Token::Colon)
                                && !self.is_at_end() && !self.check_newline() {
                                self.advance();
                            }
                        }
                        // 检查是否还有更多变量
                        if !self.match_token(&Token::Comma) {
                            break;
                        }
                    }
                    _ => break,
                }
            }
        }

        // 跳过语句分隔符
        self.match_token(&Token::Colon);
        self.skip_newlines();

        Ok(Some(Stmt::Dim {
            name,
            init,
            is_array,
            sizes,
        }))
    }

    /// 解析 Const 声明
    pub fn parse_const(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Const)?;
        let name = self.expect_ident()?;
        self.expect(Token::Eq)?;
        let value = self.parse_expr(0)?;

        // 跳过冒号（语句分隔符）
        self.match_token(&Token::Colon);
        self.skip_newlines();

        Ok(Some(Stmt::Const { name, value }))
    }

    /// 解析 Option 语句
    pub fn parse_option(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Option)?;
        self.expect_keyword(Keyword::Explicit)?;

        // 跳过冒号（语句分隔符）
        self.match_token(&Token::Colon);
        self.skip_newlines();

        Ok(Some(Stmt::OptionExplicit))
    }

    /// 解析 ReDim 语句
    pub fn parse_redim(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::ReDim)?;

        let preserve = self.match_keyword(Keyword::Preserve);
        let name = self.expect_ident()?;

        self.expect(Token::LParen)?;
        let mut sizes = vec![];
        if !self.check(&Token::RParen) {
            loop {
                sizes.push(self.parse_expr(0)?);
                if !self.match_token(&Token::Comma) {
                    break;
                }
            }
        }
        self.expect(Token::RParen)?;

        // 跳过冒号（语句分隔符）
        self.match_token(&Token::Colon);
        self.skip_newlines();

        Ok(Some(Stmt::ReDim {
            name,
            sizes,
            preserve,
        }))
    }
}
