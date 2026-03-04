//! Parser 模块 - 统一的语法解析器
//!
//! - Lexer: 手写词法分析器
//! - Parser: 统一解析器（表达式用 Pratt，语句用递归下降）
//!
//! ## 架构设计
//!
//! ```
//! parser/
//! ├── mod.rs          // 模块导出
//! ├── parser.rs       // Parser struct + 基础工具
//! ├── program.rs      // parse_program
//! ├── error.rs        // 错误定义
//! ├── expr/           // 表达式解析（Pratt 算法）
//! │   ├── mod.rs
//! │   ├── pratt.rs    // Pratt 主入口
//! │   ├── prefix.rs   // 前缀表达式
//! │   ├── infix.rs    // 中缀运算符
//! │   └── postfix.rs  // 后缀表达式
//! └── stmt/           // 语句解析（递归下降）
//!     ├── mod.rs
//!     ├── core.rs     // parse_stmt 分发
//!     ├── if_stmt.rs
//!     ├── loop_stmt.rs    // for/while/do
//!     ├── decl_stmt.rs    // dim/const
//!     ├── assign_stmt.rs  // 赋值
//!     ├── proc_stmt.rs    // function/sub
//!     └── select_stmt.rs
//! ```

mod error;
pub mod keyword;
pub mod lexer;

// Parser 核心
mod parser;
mod program;

// 表达式解析
mod expr;
mod stmt;

pub use error::ParseError;
pub use keyword::Keyword;
pub use lexer::{tokenize, Lexer, Token};
pub use parser::Parser;

/// 解析表达式（便捷函数）
pub fn parse_expr(source: &str) -> Result<crate::ast::Expr, ParseError> {
    let tokens = tokenize(source)?;
    let mut parser = Parser::new(tokens);
    parser.parse_expr(0)
}

/// 解析程序（便捷函数）
pub fn parse(source: &str) -> Result<crate::ast::Program, ParseError> {
    let tokens = tokenize(source)?;
    let mut parser = Parser::new(tokens);
    Ok(crate::ast::Program {
        statements: parser.parse_program()?,
    })
}
