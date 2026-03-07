//! 方法调用表达式求值

use crate::ast::Expr;
use crate::runtime::{BuiltinObject, RuntimeError, Value};
use crate::runtime::objects::Response;

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

        // 统一 trait dispatch：eval object 并调用方法
        let obj_val = self.eval_expr(object)?;

        if let Value::Object(mut obj) = obj_val {
            let result = obj.call_method(&method_lower, arg_values);

            // 特殊处理：Response.End 需要设置退出标志
            if let Some(response) = obj.as_any().downcast_ref::<Response>() {
                if response.is_ended() {
                    self.context.should_exit = true;
                }
            }

            return result;
        }

        Ok(Value::Empty)
    }
}
