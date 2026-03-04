//! 语句解析器（递归下降实现）
//!
//! 使用手写递归下降解析器，避免 chumsky 的泛型爆炸问题。
//! 一次只实现少量语句，确保编译时间可控。

use super::error::ParseError;
use super::keyword::Keyword;
use super::lexer::Token;
use crate::ast::{Expr, IfBranch, Param, Program, Stmt};

/// 语句解析器
pub struct StmtParser {
    tokens: Vec<Token>,
    pos: usize,
}

impl StmtParser {
    /// 创建新的语句解析器
    pub fn new(tokens: Vec<Token>) -> Self {
        StmtParser { tokens, pos: 0 }
    }

    /// 解析程序
    pub fn parse(&mut self) -> Result<Program, ParseError> {
        let mut program = Program::new();

        while !self.is_at_end() {
            // 跳过换行
            if self.check(&Token::Newline) {
                self.advance();
                continue;
            }

            // 解析语句
            if let Some(stmt) = self.parse_stmt()? {
                program.push(stmt);
            }
        }

        Ok(program)
    }

    /// 解析单条语句
    fn parse_stmt(&mut self) -> Result<Option<Stmt>, ParseError> {
        let token = self.peek()?;

        match token {
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
            Token::Keyword(Keyword::End) => Ok(None), // 由父级处理
            Token::Keyword(Keyword::Next) => Ok(None), // 由父级处理
            Token::Keyword(Keyword::Loop) => Ok(None), // 由父级处理
            Token::Keyword(Keyword::Wend) => Ok(None), // 由父级处理
            Token::Keyword(Keyword::Else) => Ok(None), // 由父级处理
            Token::Keyword(Keyword::ElseIf) => Ok(None), // 由父级处理
            Token::Eof => Ok(None),
            _ => self.parse_assignment_or_expr(),
        }
    }

    /// 解析 Dim 语句
    ///
    /// 语法：Dim name [ = expr ]
    fn parse_dim(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Dim)?;
        let name = self.expect_ident()?;
        self.skip_newlines();

        let init = if self.match_token(&Token::Eq) {
            self.skip_newlines();
            Some(self.parse_expr()?)
        } else {
            None
        };

        Ok(Some(Stmt::Dim {
            name,
            init,
            is_array: false,
            sizes: vec![],
        }))
    }

    /// 解析 Const 语句
    fn parse_const(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Const)?;
        let name = self.expect_ident()?;
        self.expect(Token::Eq)?;
        let value = self.parse_expr()?;

        Ok(Some(Stmt::Const { name, value }))
    }

    /// 解析 If 语句
    fn parse_if(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::If)?;
        let cond = self.parse_expr()?;
        self.skip_newlines();
        self.expect_keyword(Keyword::Then)?;
        self.skip_newlines();

        let mut branches = vec![IfBranch { cond, body: vec![] }];
        let mut else_block = None;

        // 解析第一个分支的 body
        loop {
            if self.check_keyword(Keyword::End) || self.check_keyword(Keyword::Else) || self.check_keyword(Keyword::ElseIf) {
                break;
            }

            if let Some(stmt) = self.parse_stmt()? {
                branches[0].body.push(stmt);
            }
            self.skip_newlines();
        }

        // 解析 ElseIf 和 Else
        while !self.check_keyword(Keyword::End) {
            if self.match_keyword(Keyword::ElseIf) {
                let cond = self.parse_expr()?;
                self.skip_newlines();
                self.expect_keyword(Keyword::Then)?;
                self.skip_newlines();

                let mut body = vec![];
                loop {
                    if self.check_keyword(Keyword::End) || self.check_keyword(Keyword::Else) || self.check_keyword(Keyword::ElseIf) {
                        break;
                    }
                    if let Some(stmt) = self.parse_stmt()? {
                        body.push(stmt);
                    }
                    self.skip_newlines();
                }
                branches.push(IfBranch { cond, body });
            } else if self.match_keyword(Keyword::Else) {
                self.skip_newlines();
                let mut body = vec![];
                loop {
                    if self.check_keyword(Keyword::End) {
                        break;
                    }
                    if let Some(stmt) = self.parse_stmt()? {
                        body.push(stmt);
                    }
                    self.skip_newlines();
                }
                else_block = Some(body);
            }
        }

        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::If)?;

        Ok(Some(Stmt::If { branches, else_block }))
    }

    /// 解析 For 循环
    fn parse_for(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::For)?;
        let var = self.expect_ident()?;
        self.expect(Token::Eq)?;
        let start = self.parse_expr()?;
        self.expect_keyword(Keyword::To)?;
        let end = self.parse_expr()?;

        let step = if self.match_keyword(Keyword::Step) {
            Some(self.parse_expr()?)
        } else {
            None
        };
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.check_keyword(Keyword::Next) {
                break;
            }
            if let Some(stmt) = self.parse_stmt()? {
                body.push(stmt);
            }
            self.skip_newlines();
        }
        self.expect_keyword(Keyword::Next)?;

        Ok(Some(Stmt::For {
            var,
            start,
            end,
            step,
            body,
        }))
    }

    /// 解析 While 循环
    fn parse_while(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::While)?;
        let cond = self.parse_expr()?;
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.check_keyword(Keyword::Wend) {
                break;
            }
            if let Some(stmt) = self.parse_stmt()? {
                body.push(stmt);
            }
            self.skip_newlines();
        }
        self.expect_keyword(Keyword::Wend)?;

        Ok(Some(Stmt::While { cond, body }))
    }

    /// 解析 Function 定义
    fn parse_function(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Function)?;
        let name = self.expect_ident()?;

        let mut params = vec![];
        if self.match_token(&Token::LParen) {
            if !self.check(&Token::RParen) {
                loop {
                    let param_name = self.expect_ident()?;
                    params.push(Param::new(param_name));
                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
            }
            self.expect(Token::RParen)?;
        }
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.check_keyword(Keyword::End) {
                let next = self.peek_ahead(1)?;
                if matches!(next, Token::Keyword(Keyword::Function)) {
                    break;
                }
            }
            if let Some(stmt) = self.parse_stmt()? {
                body.push(stmt);
            }
            self.skip_newlines();
        }

        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::Function)?;

        Ok(Some(Stmt::Function { name, params, body }))
    }

    /// 解析 Sub 定义
    fn parse_sub(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Sub)?;
        let name = self.expect_ident()?;

        let mut params = vec![];
        if self.match_token(&Token::LParen) {
            if !self.check(&Token::RParen) {
                loop {
                    let param_name = self.expect_ident()?;
                    params.push(Param::new(param_name));
                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
            }
            self.expect(Token::RParen)?;
        }
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.check_keyword(Keyword::End) {
                let next = self.peek_ahead(1)?;
                if matches!(next, Token::Keyword(Keyword::Sub)) {
                    break;
                }
            }
            if let Some(stmt) = self.parse_stmt()? {
                body.push(stmt);
            }
            self.skip_newlines();
        }

        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::Sub)?;

        Ok(Some(Stmt::Sub { name, params, body }))
    }

    /// 解析 Call 语句
    fn parse_call(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Call)?;
        let name = self.expect_ident()?;

        let mut args = vec![];
        if self.match_token(&Token::LParen) {
            if !self.check(&Token::RParen) {
                loop {
                    args.push(self.parse_expr()?);
                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
            }
            self.expect(Token::RParen)?;
        }

        Ok(Some(Stmt::Call { name, args }))
    }

    /// 解析 Set 语句
    fn parse_set(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Set)?;
        let target = self.parse_expr()?;
        self.expect(Token::Eq)?;
        let value = self.parse_expr()?;

        Ok(Some(Stmt::Set { target, value }))
    }

    /// 解析 Exit 语句
    fn parse_exit(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Exit)?;

        if self.match_keyword(Keyword::For) {
            Ok(Some(Stmt::ExitFor))
        } else if self.match_keyword(Keyword::Do) {
            Ok(Some(Stmt::ExitDo))
        } else if self.match_keyword(Keyword::Function) {
            Ok(Some(Stmt::ExitFunction))
        } else if self.match_keyword(Keyword::Sub) {
            Ok(Some(Stmt::ExitSub))
        } else {
            Err(ParseError::ParserError("Expected For, Do, Function, or Sub after Exit".to_string()))
        }
    }

    /// 解析赋值或表达式语句
    fn parse_assignment_or_expr(&mut self) -> Result<Option<Stmt>, ParseError> {
        let expr = self.parse_expr()?;

        // 检查是否是赋值
        if self.match_token(&Token::Eq) {
            let value = self.parse_expr()?;
            return Ok(Some(Stmt::Assignment { target: expr, value }));
        }

        Ok(Some(Stmt::Expr(expr)))
    }

    // ========== 表达式解析（委托给 ExprParser）==========

    fn parse_expr(&mut self) -> Result<Expr, ParseError> {
        // 使用已有的 ExprParser
        let expr_tokens = self.collect_expr_tokens()?;
        super::expr_parser::parse_expression(expr_tokens)
    }

    /// 收集表达式的 token（简化实现）
    fn collect_expr_tokens(&mut self) -> Result<Vec<Token>, ParseError> {
        // 简化：收集到语句结束符为止
        let mut tokens = vec![];

        while !self.is_at_end() && !self.is_stmt_end() {
            tokens.push(self.advance().clone());
        }

        tokens.push(Token::Eof);
        Ok(tokens)
    }

    /// 判断是否是语句结束符
    fn is_stmt_end(&self) -> bool {
        matches!(
            self.peek().unwrap(),
            Token::Newline |
            Token::Keyword(Keyword::Then) |
            Token::Keyword(Keyword::Else) |
            Token::Keyword(Keyword::ElseIf) |
            Token::Keyword(Keyword::End) |
            Token::Keyword(Keyword::Next) |
            Token::Keyword(Keyword::Loop) |
            Token::Keyword(Keyword::Wend) |
            Token::Eof
        )
    }

    // ========== 辅助方法 ==========

    fn peek(&self) -> Result<&Token, ParseError> {
        Ok(if self.pos < self.tokens.len() {
            &self.tokens[self.pos]
        } else {
            &Token::Eof
        })
    }

    fn peek_ahead(&self, offset: usize) -> Result<&Token, ParseError> {
        let pos = self.pos + offset;
        Ok(if pos < self.tokens.len() {
            &self.tokens[pos]
        } else {
            &Token::Eof
        })
    }

    fn advance(&mut self) -> &Token {
        let token = if self.pos < self.tokens.len() {
            &self.tokens[self.pos]
        } else {
            &Token::Eof
        };
        self.pos += 1;
        token
    }

    fn is_at_end(&self) -> bool {
        self.pos >= self.tokens.len() || matches!(self.tokens[self.pos], Token::Eof)
    }

    fn check(&self, token: &Token) -> bool {
        std::mem::discriminant(self.peek().unwrap()) == std::mem::discriminant(token)
    }

    fn check_keyword(&self, keyword: Keyword) -> bool {
        matches!(self.peek().unwrap(), Token::Keyword(k) if *k == keyword)
    }

    fn match_token(&mut self, token: &Token) -> bool {
        if self.check(token) {
            self.advance();
            true
        } else {
            false
        }
    }

    fn match_keyword(&mut self, keyword: Keyword) -> bool {
        if self.check_keyword(keyword) {
            self.advance();
            true
        } else {
            false
        }
    }

    fn expect(&mut self, token: Token) -> Result<&Token, ParseError> {
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

    fn expect_keyword(&mut self, keyword: Keyword) -> Result<(), ParseError> {
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

    fn expect_ident(&mut self) -> Result<String, ParseError> {
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

    fn skip_newlines(&mut self) {
        while self.match_token(&Token::Newline) {}
    }
}

/// 解析程序（便捷函数）
pub fn parse_program(tokens: Vec<Token>) -> Result<Program, ParseError> {
    let mut parser = StmtParser::new(tokens);
    parser.parse()
}
