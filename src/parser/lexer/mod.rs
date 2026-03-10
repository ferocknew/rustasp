//! Lexer 模块 - 词法分析器
//!
//! 目录结构：
//! - token.rs: Token 定义和 SpannedToken
//! - keyword.rs: VBScript 关键字定义和查找
//! - scan.rs: 扫描工具函数（scan_while, scan_until 等）
//! - lexer.rs: 主 Lexer 结构和 tokenize 逻辑

pub mod keyword;
mod lexer;
pub mod scan;
pub mod token;

// 重新导出常用类型
pub use keyword::Keyword;
pub use lexer::Lexer;
pub use token::{SpannedToken, Token};

use crate::parser::ParseError;

/// 便捷函数：对源代码进行词法分析
pub fn tokenize(source: &str) -> Result<Vec<SpannedToken>, ParseError> {
    let mut lexer = Lexer::new();
    lexer.tokenize(source)
}
