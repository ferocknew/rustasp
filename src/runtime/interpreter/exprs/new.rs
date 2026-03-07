//! New 表达式求值（类实例化）

use crate::runtime::{RuntimeError, Value};
use crate::runtime::objects::Dictionary;

use super::super::Interpreter;

impl Interpreter {
    /// 执行 New 表达式 - 创建类实例
    /// 支持内置对象和用户定义的类
    pub fn eval_new(&mut self, class_name: &str) -> Result<Value, RuntimeError> {
        let normalized_name = crate::utils::normalize_identifier(class_name);

        // 1. 先检查是否是内置对象
        if let Some(builtin) = Self::create_builtin_object(&normalized_name) {
            return Ok(builtin);
        }

        // 2. 从缓存中获取预编译的 VbsClass
        if let Some(vbs_class) = self.context.classes.get(&normalized_name) {
            // 直接使用缓存的类创建实例（避免重复构建）
            let instance = vbs_class.new_instance();
            return Ok(instance.to_value());
        }

        Err(RuntimeError::Generic(format!(
            "Class '{}' not found",
            class_name
        )))
    }

    /// 创建内置对象实例
    /// 返回 None 表示不是内置对象，需要查找用户定义的类
    fn create_builtin_object(class_name: &str) -> Option<Value> {
        match class_name {
            // Scripting.Dictionary 对象
            "dictionary" | "scripting.dictionary" => {
                Some(Value::Object(Box::new(Dictionary::new())))
            }
            // 其他内置对象可以在这里添加
            // 例如: "filesystemobject" | "scripting.filesystemobject" => ...
            _ => None,
        }
    }
}
