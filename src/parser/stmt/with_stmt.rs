//! With 语句解析

use crate::ast::Stmt;
use crate::parser::lexer::Token;
use crate::parser::Keyword;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析 With 语句
    /// 格式：With object
    ///          .member = value
    ///          .method()
    ///       End With
    pub fn parse_with(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::With)?;

        // 解析对象表达式
        let object = self.parse_expr(0)?;

        // 解析语句体，直到遇到 End With
        let body = self.parse_block_until(&[Keyword::End])?;

        // 消耗 End With
        self.expect_keyword(Keyword::End)?;
        self.skip_newlines();
        self.expect_keyword(Keyword::With)?;

        Ok(Some(Stmt::With { object, body }))
    }

    /// 解析以点号开头的语句（With 块内的成员访问）
    /// 格式：.member = value  (赋值)
    ///      .method(args)    (方法调用)
    pub fn parse_dot_stmt(&mut self) -> Result<Option<Stmt>, ParseError> {
        // 解析完整的成员访问表达式
        // 从点号开始解析，这会产生一个 GetMember 表达式
        let expr = self.parse_expr(0)?;

        // 检查是否是赋值
        // 如果后面是 =，则是赋值语句
        if self.match_token(&Token::Eq) {
            // 这是赋值语句
            let value = self.parse_expr(0)?;
            return Ok(Some(Stmt::Assignment {
                target: expr,
                value,
            }));
        }

        // 否则是表达式语句（方法调用或属性访问）
        Ok(Some(Stmt::Expr(expr)))
    }
}
