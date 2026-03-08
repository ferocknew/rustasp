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

        // 特殊处理：Err 对象方法调用
        if let Expr::Variable(name) = object {
            if name.to_lowercase() == "err" {
                return self.eval_err_method(&method_lower);
            }
        }

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

    /// 处理 Err 对象的方法调用
    fn eval_err_method(&mut self, method: &str) -> Result<Value, RuntimeError> {
        match method {
            "clear" => {
                self.context.err.clear();
                Ok(Value::Empty)
            }
            _ => Err(RuntimeError::MethodNotFound(format!("Err.{}", method))),
        }
    }
}
