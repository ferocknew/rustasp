//! 项目级公共工具模块
//!
//! 提供跨模块共享的工具函数

pub mod vbscript;

pub use vbscript::{identifier_eq, identifier_matches, normalize_identifier};
