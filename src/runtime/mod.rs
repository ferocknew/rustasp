//! Runtime 模块 - 解释执行层
//!
//! 执行 AST、管理变量作用域、管理函数调用、实现弱类型系统

mod interpreter;
mod context;
mod scope;
mod error;

pub mod value;

pub use interpreter::Interpreter;
pub use context::Context;
pub use scope::Scope;
pub use error::RuntimeError;
pub use value::Value;
