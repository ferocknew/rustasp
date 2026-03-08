//! 方法调用表达式求值

use crate::ast::Expr;
use crate::runtime::{RuntimeError, Value};

use super::super::Interpreter;

impl Interpreter {
    /// 执行方法调用
    pub fn eval_method(
        &mut self,
        object: &Expr,
        method: &str,
        args: &[Expr],
    ) -> Result<Value, RuntimeError> {
        // 计算参数
        let arg_values: Result<Vec<Value>, _> =
            args.iter().map(|e| self.eval_expr(e)).collect();
        let arg_values = arg_values?;

        let method_lower = method.to_lowercase();

        // 先计算对象表达式
        let obj_val = self.eval_expr(object)?;

        // 通过 Arc<Mutex<dyn BuiltinObject>> 调用方法
        // Arc::clone 只增加引用计数，所有引用操作同一个对象
        if let Value::Object(ref obj) = obj_val {
            let result = obj.lock()
                .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?
                .call_method(&method_lower, arg_values);

            return result;
        }

        Ok(Value::Empty)
    }
}
