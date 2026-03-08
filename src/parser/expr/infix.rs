//! 中缀运算符解析

use crate::ast::BinaryOp;
use crate::parser::Keyword;
use crate::parser::lexer::Token;
use crate::parser::{ParseError, Parser};

impl Parser {
    /// 获取中缀运算符的绑定优先级
    ///
    /// 返回 (left_bp, right_bp)
    /// - left_bp: 左侧结合强度
    /// - right_bp: 右侧结合强度
    pub(super) fn infix_binding_power(&self) -> Option<(u8, u8)> {
        match self.peek() {
            // 逻辑或 (最低优先级)
            Token::Keyword(kw) if kw.is_or() => Some((1, 2)),

            // 逻辑与
            Token::Keyword(kw) if kw.is_and() => Some((3, 4)),

            // 比较（相等性）
            Token::Eq | Token::Ne => Some((5, 6)),

            // 比较（大小）
            Token::Lt | Token::Le | Token::Gt | Token::Ge => Some((7, 8)),

            // 字符串连接
            Token::Ampersand => Some((9, 10)),

            // 加减
            Token::Plus | Token::Minus => Some((11, 12)),

            // 乘除
            Token::Star | Token::Slash | Token::Backslash => Some((13, 14)),

            // 幂运算（右结合）
            Token::Caret => Some((16, 15)),

            // Mod
            Token::Keyword(Keyword::Mod) => Some((13, 14)),

            // Is 运算符
            Token::Keyword(Keyword::Is) => Some((5, 6)),

            _ => None,
        }
    }

    /// 将 Token 转换为 BinaryOp
    pub(super) fn token_to_binary_op(&self, token: &Token) -> Result<BinaryOp, ParseError> {
        match token {
            Token::Plus => Ok(BinaryOp::Add),
            Token::Minus => Ok(BinaryOp::Sub),
            Token::Star => Ok(BinaryOp::Mul),
            Token::Slash => Ok(BinaryOp::Div),
            Token::Backslash => Ok(BinaryOp::IntDiv),
            Token::Caret => Ok(BinaryOp::Pow),
            Token::Ampersand => Ok(BinaryOp::Concat),
            Token::Eq => Ok(BinaryOp::Eq),
            Token::Ne => Ok(BinaryOp::Ne),
            Token::Lt => Ok(BinaryOp::Lt),
            Token::Le => Ok(BinaryOp::Le),
            Token::Gt => Ok(BinaryOp::Gt),
            Token::Ge => Ok(BinaryOp::Ge),
            Token::Keyword(kw) if kw.is_and() => Ok(BinaryOp::And),
            Token::Keyword(kw) if kw.is_or() => Ok(BinaryOp::Or),
            Token::Keyword(Keyword::Mod) => Ok(BinaryOp::Mod),
            Token::Keyword(Keyword::Is) => Ok(BinaryOp::Is),
            _ => Err(ParseError::ParserError(format!(
                "Expected binary operator, got {:?}",
                token
            ))),
        }
    }
}
