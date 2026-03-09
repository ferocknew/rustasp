//! 语句解析核心 - parse_stmt 分发器

use crate::ast::Stmt;
use crate::parser::Keyword;
use crate::parser::lexer::Token;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析单条语句（核心分发器）
    pub fn parse_stmt(&mut self) -> Result<Option<Stmt>, ParseError> {
        match self.peek() {
            // 声明语句
            Token::Keyword(Keyword::Dim) => self.parse_dim(),
            Token::Keyword(Keyword::Const) => self.parse_const(),
            Token::Keyword(Keyword::Option) => self.parse_option(),
            Token::Keyword(Keyword::ReDim) => self.parse_redim(),

            // 控制流语句
            Token::Keyword(Keyword::If) => self.parse_if(),
            Token::Keyword(Keyword::For) => self.parse_for(),
            Token::Keyword(Keyword::While) => self.parse_while(),
            Token::Keyword(Keyword::Select) => self.parse_select(),
            Token::Keyword(Keyword::Do) => self.parse_do(),

            // 错误处理语句
            Token::Keyword(Keyword::On) => self.parse_on_error(),

            // 类定义
            Token::Keyword(Keyword::Class) => self.parse_class(),

            // 函数/过程
            Token::Keyword(Keyword::Function) => self.parse_function(),
            Token::Keyword(Keyword::Sub) => self.parse_sub(),
            Token::Keyword(Keyword::Call) => self.parse_call(),
            Token::Keyword(Keyword::Exit) => self.parse_exit(),
            Token::Keyword(Keyword::Set) => self.parse_set(),

            // 动态执行语句
            Token::Keyword(Keyword::Execute) => self.parse_execute(),
            Token::Keyword(Keyword::ExecuteGlobal) => self.parse_execute_global(),

            // 终止符 - 返回 None
            Token::Keyword(Keyword::End)
            | Token::Keyword(Keyword::Next)
            | Token::Keyword(Keyword::Loop)
            | Token::Keyword(Keyword::Wend)
            | Token::Keyword(Keyword::Else)
            | Token::Keyword(Keyword::ElseIf)
            | Token::Keyword(Keyword::Case)
            | Token::Keyword(Keyword::Until) => Ok(None),

            // 冒号 - 空语句（用于语句分隔）
            Token::Colon => {
                self.advance(); // consume the colon
                Ok(None) // empty statement
            }

            Token::Eof => Ok(None),

            // 标识符 - 可能是赋值或表达式
            Token::Ident(_) => self.parse_ident_stmt(),

            // 其他情况当作表达式语句
            _ => self.parse_expr_stmt(),
        }
    }

    /// 解析 On Error 语句
    fn parse_on_error(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::On)?;

        // 检查是否是 Error 关键字
        if !self.match_keyword(Keyword::Error) {
            return Err(ParseError::ParserError("Expected 'Error' keyword".to_string()));
        }

        // 检查下一个关键字
        match self.peek() {
            Token::Keyword(Keyword::Resume) => {
                self.expect_keyword(Keyword::Resume)?;

                // 检查是否是 Next
                match self.peek() {
                    Token::Keyword(Keyword::Next) => {
                        self.expect_keyword(Keyword::Next)?;
                        Ok(Some(Stmt::OnErrorResumeNext))
                    }
                    _ => Err(ParseError::ParserError("Expected 'Next' keyword".to_string())),
                }
            }
            Token::Ident(ident) if ident.eq_ignore_ascii_case("goto") => {
                self.advance(); // 消耗 goto
                // 检查是否是 0
                match self.peek() {
                    Token::Number(n) if *n == 0.0 => {
                        self.advance();
                        Ok(Some(Stmt::OnErrorGoto0))
                    }
                    _ => Err(ParseError::ParserError("Expected '0'".to_string())),
                }
            }
            _ => Err(ParseError::ParserError("Expected 'Resume' or 'Goto'".to_string())),
        }
    }

    /// 解析 Execute 语句
    /// 解析 Execute 语句
    /// Execute 语句格式: Execute expression
    fn parse_execute(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Execute)?;

        // 解析表达式（通常是字符串表达式）
        let expr = self.parse_expr(0)?;

        Ok(Some(Stmt::Execute(expr)))
    }

    /// 解析 ExecuteGlobal 语句
    /// ExecuteGlobal 语句格式: ExecuteGlobal expression
    fn parse_execute_global(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::ExecuteGlobal)?;

        // 解析表达式（通常是字符串表达式）
        let expr = self.parse_expr(0)?;

        Ok(Some(Stmt::ExecuteGlobal(expr)))
    }

    /// 解析语句列表直到遇到终止关键字
    /// 用于 If、Else、While、For 等语句体
    pub fn parse_stmt_list_until(
        &mut self,
        end_keywords: &[Keyword],
    ) -> Result<Vec<Stmt>, ParseError> {
        let mut stmts = vec![];

        loop {
            // 跳过前置的换行符
            self.skip_newlines();

            // 检查是否结束
            if self.is_at_end() {
                break;
            }

            // 检查是否遇到终止关键字
            if end_keywords.iter().any(|k| self.check_keyword(*k)) {
                break;
            }

            // 解析一条语句
            if let Some(stmt) = self.parse_stmt()? {
                stmts.push(stmt);
            }

            // 如果遇到冒号，继续解析下一条语句（冒号分隔符）
            // 如果遇到换行符，结束（单行语句结束）
            // 如果遇到终止关键字，结束
            if self.match_token(&Token::Colon) {
                // 冒号分隔符，跳过连续的冒号（空语句）
                while self.match_token(&Token::Colon) {}
                continue;
            }

            // 检查是否遇到终止关键字
            if end_keywords.iter().any(|k| self.check_keyword(*k)) {
                break;
            }

            // 如果不是冒号，检查是否有换行符
            // 如果有换行符，说明语句列表结束
            if self.check(&Token::Newline) {
                break;
            }

            // 如果遇到文件结束，也结束
            if self.is_at_end() {
                break;
            }
        }

        Ok(stmts)
    }
}
