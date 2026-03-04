//! 语句解析核心 - parse_stmt 分发器

use crate::ast::Stmt;
use crate::parser::keyword::Keyword;
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

            // 函数/过程
            Token::Keyword(Keyword::Function) => self.parse_function(),
            Token::Keyword(Keyword::Sub) => self.parse_sub(),
            Token::Keyword(Keyword::Call) => self.parse_call(),
            Token::Keyword(Keyword::Exit) => self.parse_exit(),
            Token::Keyword(Keyword::Set) => self.parse_set(),

            // 终止符 - 返回 None
            Token::Keyword(Keyword::End)
            | Token::Keyword(Keyword::Next)
            | Token::Keyword(Keyword::Loop)
            | Token::Keyword(Keyword::Wend)
            | Token::Keyword(Keyword::Else)
            | Token::Keyword(Keyword::ElseIf)
            | Token::Keyword(Keyword::Case)
            | Token::Keyword(Keyword::Until) => Ok(None),

            Token::Eof => Ok(None),

            // 标识符 - 可能是赋值或表达式
            Token::Ident(_) => self.parse_ident_stmt(),

            // 其他情况当作表达式语句
            _ => self.parse_expr_stmt(),
        }
    }
}
