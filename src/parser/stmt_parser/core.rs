//! 语句解析器核心结构
//!
//! 定义 StmtParser 结构体和核心方法

use crate::ast::{Expr, Program, Stmt};
use crate::parser::error::ParseError;
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;
use std::io::Write;

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
        eprintln!("DEBUG parse_program: 开始，tokens.len()={}", self.tokens.len());
        let _ = std::io::stdout().flush();
        let mut program = Program::new();
        let mut loop_count = 0;
        let mut last_pos = self.pos;

        while !self.is_at_end() {
            loop_count += 1;
            eprintln!("DEBUG parse_program: 迭代 {}, pos={}, peek={:?}",
                loop_count, self.pos, self.peek());
            let _ = std::io::stdout().flush();

            if loop_count > 100 {
                eprintln!("DEBUG parse_program: 循环次数过多！");
                let _ = std::io::stdout().flush();
                return Err(ParseError::ParserError(format!(
                    "解析程序时检测到可能的死循环（当前 token: {:?}）",
                    self.peek()
                )));
            }

            // 检查位置是否前进
            if self.pos == last_pos && loop_count > 1 {
                eprintln!("DEBUG parse_program: 位置未前进！");
                let _ = std::io::stdout().flush();
                return Err(ParseError::ParserError(format!(
                    "解析器卡住，位置未前进（pos={}, token: {:?}）",
                    self.pos, self.peek()
                )));
            }
            last_pos = self.pos;

            // 跳过换行和冒号（语句分隔符）
            if self.check(&Token::Newline) || self.check(&Token::Colon) {
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
        // 先收集所有 token 直到语句结束
        let mut tokens = vec![];
        while !self.is_at_end() && !self.is_stmt_end() {
            tokens.push(self.advance().clone());
        }

        // 检查是否是赋值语句（查找非表达式内的 =）
        let eq_pos = find_assignment_eq(&tokens);
        if let Some(eq_pos) = eq_pos {
            // 分割为 target 和 value
            let mut target_tokens = tokens[..eq_pos].to_vec();
            target_tokens.push(Token::Eof);
            let mut value_tokens = tokens[eq_pos + 1..].to_vec();
            value_tokens.push(Token::Eof);
            
            let target = crate::parser::expr_parser::parse_expression(target_tokens)?;
            let value = crate::parser::expr_parser::parse_expression(value_tokens)?;
            return Ok(Some(Stmt::Assignment { target, value }));
        }
        
        // 检查是否是不带括号的方法调用 (obj.method arg1, arg2)
        let call_pos = find_method_call_position(&tokens);
        if let Some(call_pos) = call_pos {
            let mut obj_tokens = tokens[..call_pos].to_vec();
            obj_tokens.push(Token::Eof);
            
            let obj_expr = crate::parser::expr_parser::parse_expression(obj_tokens)?;
            
            // 检查是否是 Property（方法名）
            if let Expr::Property { object, property } = obj_expr {
                // 解析参数
                let args_tokens = tokens[call_pos..].to_vec();
                let args = parse_method_args(args_tokens)?;
                
                return Ok(Some(Stmt::Expr(Expr::Method {
                    object,
                    method: property,
                    args,
                })));
            }
        }

        // 普通表达式
        tokens.push(Token::Eof);
        let expr = crate::parser::expr_parser::parse_expression(tokens)?;
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
                | Token::Colon  // 冒号也作为语句分隔符
                | Token::Keyword(Keyword::Then)
                | Token::Keyword(Keyword::Else)
                | Token::Keyword(Keyword::ElseIf)
                | Token::Keyword(Keyword::End)
                | Token::Keyword(Keyword::Next)
                | Token::Keyword(Keyword::Loop)
                | Token::Keyword(Keyword::Wend)
                | Token::Keyword(Keyword::Case)  // Case 也作为语句分隔符
                | Token::Eof
        )
    }
}

/// 解析程序（便捷函数）
pub fn parse_program(tokens: Vec<Token>) -> Result<Program, ParseError> {
    let mut parser = StmtParser::new(tokens);
    parser.parse()
}

// ========== 辅助函数 ==========

/// 查找赋值语句中的等号位置（不在括号内的 =）
fn find_assignment_eq(tokens: &[Token]) -> Option<usize> {
    let mut paren_depth = 0;
    for (i, token) in tokens.iter().enumerate() {
        match token {
            Token::LParen => paren_depth += 1,
            Token::RParen => {
                if paren_depth > 0 {
                    paren_depth -= 1;
                }
            }
            Token::Eq if paren_depth == 0 => return Some(i),
            _ => {}
        }
    }
    None
}

/// 查找方法调用参数的起始位置
/// 方法调用模式：obj.method arg1, arg2
/// 返回参数开始的位置
fn find_method_call_position(tokens: &[Token]) -> Option<usize> {
    // 找到 . 后面的标识符位置
    let mut found_dot = false;
    for (i, token) in tokens.iter().enumerate() {
        match token {
            Token::Dot => found_dot = true,
            Token::Ident(_) if found_dot => {
                // 检查下一个 token 是否不是 ( 或 . 或操作符
                if i + 1 < tokens.len() {
                    let next = &tokens[i + 1];
                    if !matches!(
                        next,
                        Token::LParen | Token::RParen | Token::Dot | Token::Eq | Token::Newline | Token::Colon | Token::Eof
                    ) {
                        // 检查 next 是否不是操作符
                        if !is_operator(next) {
                            return Some(i + 1);
                        }
                    }
                }
                found_dot = false;
            }
            _ => found_dot = false,
        }
    }
    None
}

/// 检查 token 是否是操作符
fn is_operator(token: &Token) -> bool {
    matches!(
        token,
        Token::Plus | Token::Minus | Token::Star | Token::Slash
            | Token::Backslash | Token::Caret | Token::Ampersand
            | Token::Eq | Token::Ne | Token::Lt | Token::Le | Token::Gt | Token::Ge
            | Token::Keyword(Keyword::And) | Token::Keyword(Keyword::Or) | Token::Keyword(Keyword::Mod)
    )
}

/// 解析方法参数（逗号分隔的表达式列表）
fn parse_method_args(tokens: Vec<Token>) -> Result<Vec<Expr>, ParseError> {
    if tokens.is_empty() {
        return Ok(vec![]);
    }
    
    let mut args = vec![];
    let mut current_tokens = vec![];
    
    for token in tokens {
        if token == Token::Comma {
            if !current_tokens.is_empty() {
                current_tokens.push(Token::Eof);
                let arg = crate::parser::expr_parser::parse_expression(current_tokens)?;
                args.push(arg);
                current_tokens = vec![];
            }
        } else {
            current_tokens.push(token);
        }
    }
    
    // 处理最后一个参数
    if !current_tokens.is_empty() {
        current_tokens.push(Token::Eof);
        let arg = crate::parser::expr_parser::parse_expression(current_tokens)?;
        args.push(arg);
    }
    
    Ok(args)
}
