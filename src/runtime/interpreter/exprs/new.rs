//! New 表达式求值（类实例化）

use crate::runtime::{RuntimeError, Value};

use super::super::Interpreter;

impl Interpreter {
    /// 执行 New 表达式 - 创建类实例（使用缓存的 VbsClass）
    pub fn eval_new(&mut self, class_name: &str) -> Result<Value, RuntimeError> {
        let normalized_name = crate::utils::normalize_identifier(class_name);

        // 从缓存中获取预编译的 VbsClass
        if let Some(vbs_class) = self.context.classes.get(&normalized_name) {
            // 直接使用缓存的类创建实例（避免重复构建）
            let instance = vbs_class.new_instance();
            Ok(instance.to_value())
        } else {
            Err(RuntimeError::Generic(format!(
                "Class '{}' not found",
                class_name
            )))
        }
    }
}
