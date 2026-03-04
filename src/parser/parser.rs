//! 语法解析器

use super::error::ParseError;
use super::keyword::Keyword;
use super::lexer::Token;
use crate::ast::{BinaryOp, Expr, Program, Stmt, UnaryOp};
use chumsky::prelude::*;
use chumsky::Parser as ChumskyParser;

/// 语法解析器
pub struct Parser;

impl Parser {
    pub fn new() -> Self {
        Parser
    }

    /// 解析 Token 列表为 Program
    pub fn parse(&self, tokens: &[Token]) -> Result<Program, ParseError> {
        parse_tokens(tokens)
    }
}

impl Default for Parser {
    fn default() -> Self {
        Self::new()
    }
}

/// 解析 Token 列表为 Program
pub fn parse_tokens(tokens: &[Token]) -> Result<Program, ParseError> {
    let parser = create_parser();
    parser
        .parse(tokens.to_vec())
        .map_err(|e| ParseError::ParserError(format!("{:?}", e)))
}

/// 解析表达式
pub fn parse_expr_tokens(tokens: &[Token]) -> Result<Expr, ParseError> {
    let parser = create_expr_parser();
    parser
        .parse(tokens.to_vec())
        .map_err(|e| ParseError::ParserError(format!("{:?}", e)))
}

fn create_parser() -> impl ChumskyParser<Token, Program, Error = Simple<Token>> {
    let stmt = create_stmt_parser();
    stmt.then_ignore(just(Token::Newline).or(just(Token::Eof)).repeated())
        .repeated()
        .map(|statements| Program { statements })
}

fn create_stmt_parser() -> impl ChumskyParser<Token, Stmt, Error = Simple<Token>> {
    let expr = create_expr_parser();

    // Dim 语句
    let dim_stmt = just(Token::Keyword(Keyword::Dim))
        .ignore_then(select! { Token::Ident(name) => name })
        .then(
            just(Token::LParen)
                .ignore_then(create_expr_parser().separated_by(just(Token::Comma)))
                .then_ignore(just(Token::RParen))
                .or_not(),
        )
        .then(just(Token::Eq).ignore_then(create_expr_parser()).or_not())
        .map(|((name, sizes), init)| Stmt::Dim {
            name,
            init,
            is_array: sizes.is_some(),
            sizes: sizes.unwrap_or_default(),
        });

    // 表达式语句
    let expr_stmt = expr.map(Stmt::Expr);

    dim_stmt.or(expr_stmt)
}

fn create_expr_parser() -> impl ChumskyParser<Token, Expr, Error = Simple<Token>> {
    recursive(|expr| {
        let atom = select! {
            Token::String(s) => Expr::String(s),
            Token::Number(n) => Expr::Number(n),
            Token::Boolean(b) => Expr::Boolean(b),
            Token::Ident(name) => Expr::Variable(name),
        };

        // 成员访问和方法调用
        let member_or_call = atom
            .clone()
            .then(
                just(Token::Dot)
                    .ignore_then(select! { Token::Ident(name) => name })
                    .map(|name| Either::Left(name))
                    .or(just(Token::LParen)
                        .ignore_then(expr.clone().separated_by(just(Token::Comma)))
                        .then_ignore(just(Token::RParen))
                        .map(Either::Right))
                    .repeated(),
            )
            .map(|(first, suffixes)| {
                suffixes
                    .into_iter()
                    .fold(first, |acc, suffix| match suffix {
                        Either::Left(member) => Expr::Property {
                            object: Box::new(acc),
                            property: member,
                        },
                        Either::Right(args) => match acc {
                            Expr::Variable(name) => Expr::Call { name, args },
                            other => Expr::Method {
                                object: Box::new(other),
                                method: "call".to_string(),
                                args,
                            },
                        },
                    })
            });

        // 一元运算符
        let unary = just(Token::Minus)
            .map(|_| UnaryOp::Neg)
            .or(just(Token::Keyword(Keyword::Not)).map(|_| UnaryOp::Not))
            .then(member_or_call.clone())
            .map(|(op, operand)| Expr::Unary {
                op,
                operand: Box::new(operand),
            })
            .or(member_or_call);

        // 二元运算符
        let op = just(Token::Plus)
            .map(|_| BinaryOp::Add)
            .or(just(Token::Minus).map(|_| BinaryOp::Sub))
            .or(just(Token::Star).map(|_| BinaryOp::Mul))
            .or(just(Token::Slash).map(|_| BinaryOp::Div))
            .or(just(Token::Backslash).map(|_| BinaryOp::IntDiv))
            .or(just(Token::Ampersand).map(|_| BinaryOp::Concat))
            .or(just(Token::Eq).map(|_| BinaryOp::Eq))
            .or(just(Token::Ne).map(|_| BinaryOp::Ne))
            .or(just(Token::Lt).map(|_| BinaryOp::Lt))
            .or(just(Token::Le).map(|_| BinaryOp::Le))
            .or(just(Token::Gt).map(|_| BinaryOp::Gt))
            .or(just(Token::Ge).map(|_| BinaryOp::Ge));

        unary
            .then(op.then(unary).repeated())
            .foldl(|left, (op, right)| Expr::Binary {
                left: Box::new(left),
                op,
                right: Box::new(right),
            })
    })
}

/// Either 类型（用于解析器）
enum Either<L, R> {
    Left(L),
    Right(R),
}
