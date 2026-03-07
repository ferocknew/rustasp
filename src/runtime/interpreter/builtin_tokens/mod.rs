//! 内置函数 Token 识别模块
//!
//! 使用 Token Key 而非字符串匹配，提高内置函数调用性能
//! 将函数名映射为整数 ID，通过 match/switch 快速分发

mod token;
mod registry;
mod executor;

pub use token::BuiltinToken;
pub use registry::TokenRegistry;
pub use executor::{BuiltinExecutor, execute_builtin};
