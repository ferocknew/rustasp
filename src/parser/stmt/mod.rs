//! 语句解析模块

mod assign_stmt;
mod core;
mod decl_stmt;
mod if_stmt;
mod loop_stmt;
mod proc_stmt;
mod select_stmt;
mod with_stmt;

// 所有语句解析方法都通过 impl Parser 扩展
