//! 统一 Parser - 表达式用 Pratt，语句用递归下降
//!
//! 所有解析直接操作 self.pos，不复制 tokens

use crate::ast::{BinaryOp, CaseClause, Expr, IfBranch, Param, Stmt, UnaryOp};
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;
use crate::parser::ParseError;

/// 统一解析器
pub struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    /// 创建新的解析器
    pub fn new(tokens: Vec<Token>) -> Self {
        Parser { tokens, pos: 0 }
    }

    // ==================== 程序入口 ====================

    /// 解析程序
    pub fn parse_program(&mut self) -> Result<Vec<Stmt>, ParseError> {
        let mut stmts = vec![];

        while !self.is_at_end() {
            self.skip_newlines();

            if self.is_at_end() {
                break;
            }

            if let Some(stmt) = self.parse_stmt()? {
                stmts.push(stmt);
            }
        }

        Ok(stmts)
    }

    // ==================== 语句解析 ====================

    /// 解析单条语句
    fn parse_stmt(&mut self) -> Result<Option<Stmt>, ParseError> {
        match self.peek() {
            Token::Keyword(Keyword::Dim) => self.parse_dim(),
            Token::Keyword(Keyword::Const) => self.parse_const(),
            Token::Keyword(Keyword::Option) => self.parse_option(),
            Token::Keyword(Keyword::If) => self.parse_if(),
            Token::Keyword(Keyword::For) => self.parse_for(),
            Token::Keyword(Keyword::While) => self.parse_while(),
            Token::Keyword(Keyword::Select) => self.parse_select(),
            Token::Keyword(Keyword::Function) => self.parse_function(),
            Token::Keyword(Keyword::Sub) => self.parse_sub(),
            Token::Keyword(Keyword::Call) => self.parse_call(),
            Token::Keyword(Keyword::Set) => self.parse_set(),
            Token::Keyword(Keyword::Exit) => self.parse_exit(),
            Token::Keyword(Keyword::ReDim) => self.parse_redim(),

            // 终止符 - 返回 None
            Token::Keyword(Keyword::End)
            | Token::Keyword(Keyword::Next)
            | Token::Keyword(Keyword::Loop)
            | Token::Keyword(Keyword::Wend)
            | Token::Keyword(Keyword::Else)
            | Token::Keyword(Keyword::ElseIf)
            | Token::Keyword(Keyword::Case) => Ok(None),

            Token::Eof => Ok(None),

            // 标识符 - 可能是赋值或表达式
            Token::Ident(_) => self.parse_ident_stmt(),

            // 其他情况当作表达式语句
            _ => {
                let expr = self.parse_expr(0)?;
                self.skip_newlines();
                Ok(Some(Stmt::Expr(expr)))
            }
        }
    }

    /// 解析标识符开头的语句（赋值或表达式）
    fn parse_ident_stmt(&mut self) -> Result<Option<Stmt>, ParseError> {
        // 解析左侧表达式
        let lhs = self.parse_expr(0)?;

        // 检查是否是赋值
        if self.check(&Token::Eq) {
            self.advance(); // 消耗 =
            let value = self.parse_expr(0)?;
            self.skip_newlines();
            return Ok(Some(Stmt::Assignment {
                target: lhs,
                value,
            }));
        }

        // 否则是表达式语句
        self.skip_newlines();
        Ok(Some(Stmt::Expr(lhs)))
    }

    // ==================== 控制流语句 ====================

    /// 解析 If 语句
    fn parse_if(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::If)?;
        let cond = self.parse_expr(0)?;
        self.skip_newlines();
        self.expect_keyword(Keyword::Then)?;
        self.skip_newlines();

        // 判断是单行 If 还是多行 If
        // 单行 If: if x = 1 then Response.Write("Yes")
        // 多行 If: if x = 1 then \n ... \n end if

        // 检查是否是单行 If（下一个不是 End/Else/ElseIf 且有内容）
        if !self.check_keyword(Keyword::End)
            && !self.check_keyword(Keyword::Else)
            && !self.check_keyword(Keyword::ElseIf)
            && !self.is_at_end()
        {
            // 检查下一个 token 是否在当前行（非换行）
            // 如果遇到换行，说明是多行 If
            if self.check(&Token::Newline) {
                // 多行 If
                return self.parse_if_block(cond);
            }

            // 单行 If - 只解析一条语句
            let stmt = self.parse_stmt()?;
            let body = stmt.map_or(vec![], |s| vec![s]);

            // 检查是否有 Else
            let else_block = if self.match_keyword(Keyword::Else) {
                let else_stmt = self.parse_stmt()?;
                Some(else_stmt.map_or(vec![], |s| vec![s]))
            } else {
                None
            };

            return Ok(Some(Stmt::If {
                branches: vec![IfBranch { cond, body }],
                else_block,
            }));
        }

        // 多行 If
        self.parse_if_block(cond)
    }

    /// 解析多行 If 块
    fn parse_if_block(&mut self, first_cond: Expr) -> Result<Option<Stmt>, ParseError> {
        let mut branches = vec![IfBranch {
            cond: first_cond,
            body: vec![],
        }];
        let mut else_block = None;

        // 解析第一个 If 分支的 body
        loop {
            if self.is_at_end()
                || self.check_keyword(Keyword::End)
                || self.check_keyword(Keyword::Else)
                || self.check_keyword(Keyword::ElseIf)
            {
                break;
            }

            if let Some(stmt) = self.parse_stmt()? {
                branches[0].body.push(stmt);
            }
            self.skip_newlines();
        }

        // 解析 ElseIf 和 Else 分支
        while !self.check_keyword(Keyword::End) {
            if self.match_keyword(Keyword::ElseIf) {
                let cond = self.parse_expr(0)?;
                self.skip_newlines();
                self.expect_keyword(Keyword::Then)?;
                self.skip_newlines();

                let mut body = vec![];
                loop {
                    if self.is_at_end()
                        || self.check_keyword(Keyword::End)
                        || self.check_keyword(Keyword::Else)
                        || self.check_keyword(Keyword::ElseIf)
                    {
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
                    if self.is_at_end() || self.check_keyword(Keyword::End) {
                        break;
                    }
                    if let Some(stmt) = self.parse_stmt()? {
                        body.push(stmt);
                    }
                    self.skip_newlines();
                }
                else_block = Some(body);
            } else {
                break;
            }
        }

        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::If)?;

        Ok(Some(Stmt::If {
            branches,
            else_block,
        }))
    }

    /// 解析 For 循环
    fn parse_for(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::For)?;
        let var = self.expect_ident()?;
        self.expect(&Token::Eq)?;
        let start = self.parse_expr(0)?;
        self.expect_keyword(Keyword::To)?;
        let end = self.parse_expr(0)?;

        let step = if self.match_keyword(Keyword::Step) {
            Some(self.parse_expr(0)?)
        } else {
            None
        };
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.is_at_end() || self.check_keyword(Keyword::Next) {
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
        let cond = self.parse_expr(0)?;
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.is_at_end() || self.check_keyword(Keyword::Wend) {
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

    /// 解析 Select Case 语句
    fn parse_select(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Select)?;
        self.expect_keyword(Keyword::Case)?;

        let expr = self.parse_expr(0)?;
        self.skip_newlines();

        let mut cases = vec![];
        let mut else_block = None;

        loop {
            if self.is_at_end() || self.check_keyword(Keyword::End) {
                break;
            }

            self.expect_keyword(Keyword::Case)?;

            if self.match_keyword(Keyword::Else) {
                self.skip_newlines();
                let mut body = vec![];
                loop {
                    if self.is_at_end()
                        || self.check_keyword(Keyword::End)
                        || self.check_keyword(Keyword::Case)
                    {
                        break;
                    }
                    if let Some(stmt) = self.parse_stmt()? {
                        body.push(stmt);
                    }
                    self.skip_newlines();
                }
                else_block = Some(body);
            } else {
                // 解析 Case 值列表
                let mut values = vec![];
                loop {
                    let value = self.parse_expr(0)?;
                    values.push(value);

                    if !self.match_token(&Token::Comma) {
                        break;
                    }
                }
                self.skip_newlines();

                let mut body = vec![];
                loop {
                    if self.is_at_end()
                        || self.check_keyword(Keyword::End)
                        || self.check_keyword(Keyword::Case)
                    {
                        break;
                    }
                    if let Some(stmt) = self.parse_stmt()? {
                        body.push(stmt);
                    }
                    self.skip_newlines();
                }

                cases.push(CaseClause {
                    values: Some(values),
                    body,
                });
            }
        }

        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::Select)?;

        Ok(Some(Stmt::Select {
            expr,
            cases,
            else_block,
        }))
    }

    // ==================== 声明语句 ====================

    /// 解析 Dim 声明
    fn parse_dim(&mut self) -> Result<Option<Stmt>, ParseError> {
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
            self.expect(&Token::RParen)?;
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
    fn parse_const(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Const)?;
        let name = self.expect_ident()?;
        self.expect(&Token::Eq)?;
        let value = self.parse_expr(0)?;
        self.skip_newlines();

        Ok(Some(Stmt::Const { name, value }))
    }

    /// 解析 Option 语句
    fn parse_option(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Option)?;
        self.expect_keyword(Keyword::Explicit)?;
        self.skip_newlines();

        Ok(Some(Stmt::OptionExplicit))
    }

    /// 解析 ReDim 语句
    fn parse_redim(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::ReDim)?;

        let preserve = self.match_keyword(Keyword::Preserve);
        let name = self.expect_ident()?;

        self.expect(&Token::LParen)?;
        let mut sizes = vec![];
        if !self.check(&Token::RParen) {
            loop {
                sizes.push(self.parse_expr(0)?);
                if !self.match_token(&Token::Comma) {
                    break;
                }
            }
        }
        self.expect(&Token::RParen)?;
        self.skip_newlines();

        Ok(Some(Stmt::ReDim {
            name,
            sizes,
            preserve,
        }))
    }

    // ==================== 函数/过程 ====================

    /// 解析 Function 定义
    fn parse_function(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Function)?;
        let name = self.expect_ident()?;

        let params = self.parse_params()?;
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.is_at_end() || self.check_keyword(Keyword::End) {
                break;
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

        let params = self.parse_params()?;
        self.skip_newlines();

        let mut body = vec![];
        loop {
            if self.is_at_end() || self.check_keyword(Keyword::End) {
                break;
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

    /// 解析参数列表
    fn parse_params(&mut self) -> Result<Vec<Param>, ParseError> {
        let mut params = vec![];

        if self.match_token(&Token::LParen) {
            if !self.check(&Token::RParen) {
                loop {
                    let is_byref = self.match_keyword(Keyword::ByRef);
                    let _ = self.match_keyword(Keyword::ByVal); // ByVal 是默认的

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
            self.expect(&Token::RParen)?;
        }

        Ok(params)
    }

    // ==================== 其他语句 ====================

    /// 解析 Call 语句
    fn parse_call(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Call)?;
        let name = self.expect_ident()?;

        let args = if self.match_token(&Token::LParen) {
            let args = self.parse_args()?;
            self.expect(&Token::RParen)?;
            args
        } else {
            vec![]
        };

        self.skip_newlines();
        Ok(Some(Stmt::Call { name, args }))
    }

    /// 解析 Set 语句
    fn parse_set(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Set)?;
        let target = self.parse_expr(0)?;
        self.expect(&Token::Eq)?;
        let value = self.parse_expr(0)?;
        self.skip_newlines();

        Ok(Some(Stmt::Set { target, value }))
    }

    /// 解析 Exit 语句
    fn parse_exit(&mut self) -> Result<Option<Stmt>, ParseError> {
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

    // ==================== 辅助方法 ====================

    fn peek(&self) -> &Token {
        if self.pos < self.tokens.len() {
            &self.tokens[self.pos]
        } else {
            &Token::Eof
        }
    }

    fn peek_ahead(&self, offset: usize) -> &Token {
        let pos = self.pos + offset;
        if pos < self.tokens.len() {
            &self.tokens[pos]
        } else {
            &Token::Eof
        }
    }

    fn advance(&mut self) -> &Token {
        let token = self.peek();
        self.pos += 1;
        token
    }

    fn is_at_end(&self) -> bool {
        self.pos >= self.tokens.len() || matches!(self.tokens[self.pos], Token::Eof)
    }

    fn check(&self, token: &Token) -> bool {
        std::mem::discriminant(self.peek()) == std::mem::discriminant(token)
    }

    fn check_keyword(&self, keyword: Keyword) -> bool {
        matches!(self.peek(), Token::Keyword(k) if *k == keyword)
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
                self.peek()
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
                self.peek()
            )))
        }
    }

    fn expect_ident(&mut self) -> Result<String, ParseError> {
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

    fn skip_newlines(&mut self) {
        while self.match_token(&Token::Newline) {}
    }
}
