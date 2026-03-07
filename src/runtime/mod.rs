//! Runtime 模块 - 解释执行层
//!
//! 执行 AST、管理变量作用域、管理函数调用、实现弱类型系统

mod context;
mod error;
mod interpreter;
mod scope;

pub mod value;

pub use context::{ClassDef, Context, Function};
pub use error::RuntimeError;
pub use interpreter::Interpreter;
pub use scope::Scope;
pub use value::{Value, ValueCompare, ValueConversion, ValueOps};

/// 内置对象 trait
pub trait BuiltinObject: Send + Sync {
    /// 获取属性
    fn get_property(&self, name: &str) -> Result<Value, RuntimeError>;

    /// 设置属性
    fn set_property(&mut self, name: &str, value: Value) -> Result<(), RuntimeError>;

    /// 调用方法
    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError>;

    /// 索引访问（如 Session("key")）
    fn index(&self, key: &Value) -> Result<Value, RuntimeError> {
        let key_str = ValueConversion::to_string(key);
        // 默认实现：尝试作为属性访问
        self.get_property(&key_str)
    }
}
