//! 函数和过程语句解析 - Function / Sub / Call / Exit

use crate::ast::Param;
use crate::ast::Stmt;
use crate::parser::lexer::Token;
use crate::parser::Keyword;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析 Function 定义
    pub fn parse_function(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.parse_proc(true)
    }

    /// 解析 Sub 定义
    pub fn parse_sub(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.parse_proc(false)
    }

    /// 解析过程（Function 或 Sub）定义
    fn parse_proc(&mut self, is_function: bool) -> Result<Option<Stmt>, ParseError> {
        // 期望 Function 或 Sub 关键字
        if is_function {
            self.expect_keyword(Keyword::Function)?;
        } else {
            self.expect_keyword(Keyword::Sub)?;
        }

        // 允许使用关键字作为函数名（例如 Error）
        let name = self.expect_ident_or_keyword()?;
        let params = self.parse_params()?;
        self.skip_newlines();

        // 解析函数体 - 使用空的 end_keywords，因为我们在 parse_proc 中手动处理 End Function/Sub
        let body = self.parse_proc_block()?;

        // 期望 End Function / End Sub
        self.expect_keyword(Keyword::End)?;
        if is_function {
            self.expect_keyword(Keyword::Function)?;
            Ok(Some(Stmt::Function { name, params, body }))
        } else {
            self.expect_keyword(Keyword::Sub)?;
            Ok(Some(Stmt::Sub { name, params, body }))
        }
    }

    /// 解析过程体（直到遇到 End Function / End Sub）
    fn parse_proc_block(&mut self) -> Result<Vec<Stmt>, ParseError> {
        // 使用空的 end_keywords，因为我们在 parse_proc 中手动处理 End Function/Sub
        // parse_block_until 会在遇到 End 时停止（不区分 End Function/Sub/If/Select）
        self.parse_block_until(&[])
    }

    /// 解析参数列表
    pub fn parse_params(&mut self) -> Result<Vec<Param>, ParseError> {
        let mut params = vec![];

        if self.match_token(&Token::LParen) {
            if !self.check(&Token::RParen) {
                loop {
                    // VBScript 默认是 ByRef，只有显式指定 ByVal 才是按值传递
                    let is_byref = if self.match_keyword(Keyword::ByVal) {
                        false
                    } else {
                        self.match_keyword(Keyword::ByRef); // ByRef 是默认的
                        true
                    };

                    let name = self.expect_ident()?;

                    let default = if self.match_token(&Token::Eq) {
                        Some(self.parse_expr(0)?)
                    } else {
                        None
                    };

                    params.push(Param {
                        name,
                        is_byref,
                        default,
                    });

                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
            }
            self.expect(Token::RParen)?;
        }

        Ok(params)
    }

    /// 解析 Call 语句
    pub fn parse_call(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Call)?;
        let name = self.expect_ident()?;

        let args = if self.match_token(&Token::LParen) {
            let args = self.parse_args()?;
            self.expect(Token::RParen)?;
            args
        } else {
            vec![]
        };

        self.skip_newlines();
        Ok(Some(Stmt::Call { name, args }))
    }

    /// 解析 Exit 语句
    pub fn parse_exit(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Exit)?;

        let stmt = if self.match_keyword(Keyword::For) {
            Stmt::ExitFor
        } else if self.match_keyword(Keyword::Do) {
            Stmt::ExitDo
        } else if self.match_keyword(Keyword::Function) {
            Stmt::ExitFunction
        } else if self.match_keyword(Keyword::Sub) {
            Stmt::ExitSub
        } else if self.match_keyword(Keyword::Property) {
            Stmt::ExitProperty
        } else {
            return Err(ParseError::ParserError(format!(
                "Expected For, Do, Function, Sub, or Property after Exit, got {:?}",
                self.peek()
            )));
        };

        self.skip_newlines();
        Ok(Some(stmt))
    }
}
