//! 语句解析器核心结构
//!
//! 定义 StmtParser 结构体和核心方法

use crate::ast::{Expr, Program, Stmt};
use crate::parser::error::ParseError;
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;

/// 语句解析器
pub struct StmtParser {
    pub(super) tokens: Vec<Token>,
    pub(super) pos: usize,
    /// 待插入的多变量声明
    pub(super) pending_dims: Vec<Stmt>,
}

impl StmtParser {
    /// 创建新的语句解析器
    pub fn new(tokens: Vec<Token>) -> Self {
        StmtParser {
            tokens,
            pos: 0,
            pending_dims: vec![],
        }
    }

    /// 解析程序
    pub fn parse(&mut self) -> Result<Program, ParseError> {
        let mut program = Program::new();

        while !self.is_at_end() {
            if self.check(&Token::Newline) {
                self.advance();
                continue;
            }

            while let Some(stmt) = self.pending_dims.pop() {
                program.push(stmt);
            }

            if let Some(stmt) = self.parse_stmt()? {
                program.push(stmt);
            }
        }

        while let Some(stmt) = self.pending_dims.pop() {
            program.push(stmt);
        }

        Ok(program)
    }

    /// 解析单条语句
    pub(super) fn parse_stmt(&mut self) -> Result<Option<Stmt>, ParseError> {
        match self.peek()? {
            Token::Keyword(Keyword::Dim) => self.parse_dim(),
            Token::Keyword(Keyword::Const) => self.parse_const(),
            Token::Keyword(Keyword::If) => self.parse_if(),
            Token::Keyword(Keyword::For) => self.parse_for(),
            Token::Keyword(Keyword::While) => self.parse_while(),
            Token::Keyword(Keyword::Function) => self.parse_function(),
            Token::Keyword(Keyword::Sub) => self.parse_sub(),
            Token::Keyword(Keyword::Call) => self.parse_call(),
            Token::Keyword(Keyword::Set) => self.parse_set(),
            Token::Keyword(Keyword::Exit) => self.parse_exit(),
            Token::Keyword(Keyword::End) => Ok(None),
            Token::Keyword(Keyword::Next) => Ok(None),
            Token::Keyword(Keyword::Loop) => Ok(None),
            Token::Keyword(Keyword::Wend) => Ok(None),
            Token::Keyword(Keyword::Else) => Ok(None),
            Token::Keyword(Keyword::ElseIf) => Ok(None),
            Token::Eof => Ok(None),
            _ => self.parse_assignment_or_expr(),
        }
    }

    fn parse_assignment_or_expr(&mut self) -> Result<Option<Stmt>, ParseError> {
        let expr = self.parse_expr()?;
        if self.match_token(&Token::Eq) {
            let value = self.parse_expr()?;
            return Ok(Some(Stmt::Assignment { target: expr, value }));
        }
        Ok(Some(Stmt::Expr(expr)))
    }

    /// 解析表达式
    pub(super) fn parse_expr(&mut self) -> Result<Expr, ParseError> {
        let mut tokens = vec![];
        while !self.is_at_end() && !self.is_stmt_end() {
            tokens.push(self.advance().clone());
        }
        tokens.push(Token::Eof);
        crate::parser::expr_parser::parse_expression(tokens)
    }

    fn is_stmt_end(&self) -> bool {
        matches!(
            self.peek().unwrap(),
            Token::Newline
                | Token::Keyword(Keyword::Then)
                | Token::Keyword(Keyword::Else)
                | Token::Keyword(Keyword::ElseIf)
                | Token::Keyword(Keyword::End)
                | Token::Keyword(Keyword::Next)
                | Token::Keyword(Keyword::Loop)
                | Token::Keyword(Keyword::Wend)
                | Token::Eof
        )
    }
}

/// 解析程序（便捷函数）
pub fn parse_program(tokens: Vec<Token>) -> Result<Program, ParseError> {
    let mut parser = StmtParser::new(tokens);
    parser.parse()
}
