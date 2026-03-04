//! VBScript 语法解析器

use chumsky::prelude::*;
use serde::{Deserialize, Serialize};

use super::lexer::{Keyword, Token};

/// AST 节点
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Ast {
    // 语句
    Dim { name: String, init: Option<Box<Ast>> },
    Const { name: String, value: Box<Ast> },
    Assignment { target: Box<Ast>, value: Box<Ast> },
    If { cond: Box<Ast>, then_block: Vec<Ast>, else_block: Option<Vec<Ast>> },
    For { var: String, start: Box<Ast>, end: Box<Ast>, step: Option<Box<Ast>>, body: Vec<Ast> },
    While { cond: Box<Ast>, body: Vec<Ast> },
    DoWhile { cond: Box<Ast>, body: Vec<Ast> },
    DoUntil { cond: Box<Ast>, body: Vec<Ast> },
    ExitFor,
    ExitDo,
    ExitFunction,
    ExitSub,
    Sub { name: String, params: Vec<String>, body: Vec<Ast> },
    Function { name: String, params: Vec<String>, body: Vec<Ast> },
    Call { name: String, args: Vec<Ast> },
    Class { name: String, members: Vec<Ast> },
    PropertyGet { name: String, body: Vec<Ast> },
    PropertyLet { name: String, body: Vec<Ast> },
    ReDim { name: String, sizes: Vec<Ast>, preserve: bool },
    Erase { name: String },
    Execute { code: String },
    Eval { code: String },
    OptionExplicit,
    OnErrorResumeNext,
    ResumeNext,

    // 表达式
    BinaryOp { op: BinOp, left: Box<Ast>, right: Box<Ast> },
    UnaryOp { op: UnaryOp, operand: Box<Ast> },
    CallExpr { func: Box<Ast>, args: Vec<Ast> },
    Index { object: Box<Ast>, index: Box<Ast> },
    Member { object: Box<Ast>, member: String },
    Ident(String),
    String(String),
    Number(f64),
    Boolean(bool),
    Nothing,
    Empty,
    Null,
    Array(Vec<Ast>),
    New(String),
}

/// 二元运算符
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum BinOp {
    Add, Sub, Mul, Div, IntDiv, Mod, Pow,
    Concat,
    Eq, Ne, Lt, Le, Gt, Ge,
    And, Or, Xor, Is,
}

/// 一元运算符
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum UnaryOp {
    Neg, Not,
}

/// 语法解析器
pub type Parser = impl Parser<Token, Vec<Ast>, Error = Simple<Token>>;

/// 创建语法解析器
pub fn parser() -> Parser {
    let expr = recursive(|expr| {
        let atom = select! {
            Token::String(s) => Ast::String(s),
            Token::Number(n) => Ast::Number(n),
            Token::Boolean(b) => Ast::Boolean(b),
            Token::Ident(name) => Ast::Ident(name),
        };

        let member_or_call = atom.clone()
            .then(
                just(Token::Dot)
                    .ignore_then(select! { Token::Ident(name) => name })
                    .or(just(Token::LParen)
                        .ignore_then(expr.clone().separated_by(just(Token::Comma)))
                        .then_ignore(just(Token::RParen)))
                    .repeated()
            ).map(|(first, suffixes)| {
                suffixes.into_iter().fold(first, |acc, suffix| {
                    match suffix {
                        chumsky::primitive::Either::Left(name) => Ast::Member { object: Box::new(acc), member: name },
                        chumsky::primitive::Either::Right(args) => Ast::CallExpr { func: Box::new(acc), args },
                    }
                })
            });

        member_or_call
    });

    // 简化版本：只解析表达式语句
    expr.then_ignore(just(Token::Newline).or(just(Token::Eof)))
        .repeated()
}

/// 解析 Token 列表为 AST
pub fn parse(tokens: Vec<Token>) -> Result<Vec<Ast>, String> {
    parser()
        .parse(tokens)
        .map_err(|e| format!("Parser error: {:?}", e))
}

#[cfg(test)]
mod tests {
    use super::*;
    use super::super::lexer::tokenize;

    #[test]
    fn test_parse_string() {
        let tokens = tokenize(r#""Hello""#).unwrap();
        let ast = parse(tokens).unwrap();
        assert_eq!(ast, vec![Ast::String("Hello".to_string())]);
    }

    #[test]
    fn test_parse_ident() {
        let tokens = tokenize("Response").unwrap();
        let ast = parse(tokens).unwrap();
        assert_eq!(ast, vec![Ast::Ident("Response".to_string())]);
    }
}
