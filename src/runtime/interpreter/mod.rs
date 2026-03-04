//! 解释器模块
//!
//! 将解释器拆分为多个子模块以提高可维护性

mod builtins;
mod exprs;
mod stmts;

// 重导出
pub use builtins::call_builtin_function;
