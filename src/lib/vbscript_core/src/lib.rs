//! VBScript 语言核心模块
//!
//! 提供 VBScript 的词法分析和语法解析功能

pub mod token_list;
pub mod lexer;
pub mod parser;
pub mod runtime;

pub use token_list::{Token, Keyword};
pub use lexer::Lexer;
pub use parser::Parser;
pub use runtime::{Interpreter, RuntimeError, BuiltinObject};

// 重导出 AST 类型
pub use vbscript_ast::{Expr, Stmt, Program, BinaryOp, UnaryOp, IfBranch, Param, ClassMember};
