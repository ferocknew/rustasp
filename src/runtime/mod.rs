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
// 重新导出 Rc 供其他模块使用
pub use std::rc::Rc;
pub use error::{ControlFlow, RuntimeError};
pub use interpreter::Interpreter;
pub use scope::Scope;
pub use value::{Value, ValueCompare, ValueConversion, ValueIndex, ValueOps};

use std::sync::{Arc, Mutex};

/// 数组引用类型（共享单例）
///
/// 使用 Arc<Mutex<VbsArray>> 实现共享所有权：
/// - Arc：允许多个 Value 引用同一个数组（引用计数共享）
/// - Mutex：提供内部可变性（数组修改需要 &mut self）
/// - VbsArray：支持多维数组的扁平存储
///
/// 这样数组赋值会共享同一个数组，而不是复制：
/// ```vbscript
/// arr1 = Array(1, 2, 3)
/// arr2 = arr1  ' arr2 指向同一个数组
/// arr2(0) = 99
/// Response.Write arr1(0)  ' 输出 99
/// ```
pub type ArrayRef = Arc<Mutex<VbsArray>>;

/// 内置对象引用类型（共享单例）
///
/// 使用 Arc<Mutex<>> 实现共享所有权：
/// - Arc：允许多个 Value 引用同一个对象（引用计数共享）
/// - Mutex：提供内部可变性（call_method 需要 &mut self）
///
/// 这样 Response.Write 和 context.response 操作的是同一个对象。
pub type ObjectRef = Arc<Mutex<dyn BuiltinObject>>;

/// 内置对象 trait
pub trait BuiltinObject: Send + Sync + std::fmt::Debug {
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

    /// 索引赋值（如 Session("key") = value）
    fn set_index(&mut self, key: &Value, value: Value) -> Result<(), RuntimeError> {
        let key_str = ValueConversion::to_string(key);
        // 默认实现：尝试作为属性设置
        self.set_property(&key_str, value)
    }

    /// 向下转型支持
    fn as_any(&self) -> &dyn std::any::Any;

    /// 向下转型支持（可变）
    fn as_any_mut(&mut self) -> &mut dyn std::any::Any;
}
