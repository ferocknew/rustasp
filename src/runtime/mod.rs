//! Runtime 模块 - 解释执行层
//!
//! 执行 AST、管理变量作用域、管理函数调用、实现弱类型系统

mod class;
mod context;
mod error;
mod interpreter;
mod scope;

pub mod builtins;
pub mod value;
pub mod objects;

pub use class::{VbsClass, VbsInstance};
pub use context::{ClassDef, Context, Function};
pub use error::{ControlFlow, RuntimeError};
pub use interpreter::Interpreter;
pub use scope::Scope;
pub use value::{Value, ValueCompare, ValueConversion, ValueOps};

/// 内置对象 trait
pub trait BuiltinObject: Send + Sync + std::fmt::Debug {
    /// 克隆对象
    fn clone_box(&self) -> Box<dyn BuiltinObject>;

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

    /// 向下转型支持
    fn as_any(&self) -> &dyn std::any::Any;

    /// 向下转型支持（可变）
    fn as_any_mut(&mut self) -> &mut dyn std::any::Any;
}

impl Clone for Box<dyn BuiltinObject> {
    fn clone(&self) -> Self {
        self.clone_box()
    }
}
