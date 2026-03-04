//! VBScript 语法解析器

use chumsky::prelude::*;
use vbscript_ast::{Expr, Stmt, Program, BinaryOp, UnaryOp};

use super::token_list::{Keyword, Token};

/// 语法解析器
pub type Parser = impl Parser<Token, Program, Error = Simple<Token>>;

/// 创建语法解析器
pub fn parser() -> Parser {
    // 表达式
    let expr = recursive(|expr| {
        let atom = select! {
            Token::String(s) => Expr::String(s),
            Token::Number(n) => Expr::Number(n),
            Token::Boolean(b) => Expr::Boolean(b),
            Token::Ident(name) => Expr::Variable(name),
        };

        let member_or_call = atom.clone()
            .then(
                just(Token::Dot)
                    .ignore_then(select! { Token::Ident(name) => name })
                    .map(|name| Either::Left(name))
                    .or(just(Token::LParen)
                        .ignore_then(expr.clone().separated_by(just(Token::Comma)))
                        .then_ignore(just(Token::RParen))
                        .map(Either::Right))
                    .repeated()
            ).map(|(first, suffixes)| {
                suffixes.into_iter().fold(first, |acc, suffix| {
                    match suffix {
                        Either::Left(member) => Expr::Property {
                            object: Box::new(acc),
                            property: member,
                        },
                        Either::Right(args) => Expr::Call {
                            name: match acc {
                                Expr::Variable(name) => name,
                                other => format!("{:?}", other),
                            },
                            args,
                        },
                    }
                })
            });

        // 二元运算符优先级
        let op = just(Token::Plus).map(|_| BinaryOp::Add)
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

        member_or_call
            .then(op.then(member_or_call).repeated())
            .foldl(|left, (op, right)| Expr::Binary {
                left: Box::new(left),
                op,
                right: Box::new(right),
            })
    });

    // 语句
    let stmt = just(Token::Keyword(Keyword::Dim))
        .ignore_then(select! { Token::Ident(name) => name })
        .then(
            just(Token::Eq)
                .ignore_then(expr.clone())
                .or_not()
        )
        .map(|(name, init)| Stmt::Dim {
            name,
            init,
            is_array: false,
            sizes: vec![],
        });

    let stmt = stmt
        .or(expr.clone().map(Stmt::Expr));

    let program = stmt
        .then_ignore(just(Token::Newline).or(just(Token::Eof)))
        .repeated()
        .map(|statements| Program { statements });

    program
}

/// 解析 Token 列表为 AST
pub fn parse(tokens: Vec<Token>) -> Result<Program, String> {
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
        assert_eq!(ast.statements, vec![Stmt::Expr(Expr::String("Hello".to_string()))]);
    }

    #[test]
    fn test_parse_number() {
        let tokens = tokenize("42").unwrap();
        let ast = parse(tokens).unwrap();
        assert_eq!(ast.statements, vec![Stmt::Expr(Expr::Number(42.0))]);
    }

    #[test]
    fn test_parse_variable() {
        let tokens = tokenize("Response").unwrap();
        let ast = parse(tokens).unwrap();
        assert_eq!(ast.statements, vec![Stmt::Expr(Expr::Variable("Response".to_string()))]);
    }
}
