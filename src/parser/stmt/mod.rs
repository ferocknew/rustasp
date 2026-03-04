//! 语句解析模块

mod core;
mod if_stmt;
mod loop_stmt;
mod decl_stmt;
mod assign_stmt;
mod proc_stmt;
mod select_stmt;

// 所有语句解析方法都通过 impl Parser 扩展
