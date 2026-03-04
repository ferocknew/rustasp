//! 声明语句解析 - Dim / Const / Option / ReDim

use crate::ast::{Expr, Stmt};
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;
use crate::parser::{ParseError, Parser};

impl Parser {
    /// 解析 Dim 声明
    pub fn parse_dim(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Dim)?;
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
        let init = if self.match_token(&Token::Eq) {
            Some(self.parse_expr(0)?)
        } else {
            None
        };

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
        self.skip_newlines();

        Ok(Some(Stmt::Const { name, value }))
    }

    /// 解析 Option 语句
    pub fn parse_option(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Option)?;
        self.expect_keyword(Keyword::Explicit)?;
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
        self.skip_newlines();

        Ok(Some(Stmt::ReDim {
            name,
            sizes,
            preserve,
        }))
    }
}
