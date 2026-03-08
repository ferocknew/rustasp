//! 方法调用表达式求值

use crate::ast::Expr;
use crate::runtime::{RuntimeError, Value, BuiltinObject};
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

        // 先计算对象表达式以获取对象名称
        let obj_val = self.eval_expr(object)?;

        // 特殊处理：如果是 Response 对象，直接操作 Context 中的原始对象
        // 避免克隆导致的修改丢失问题
        if let Value::Object(ref obj) = obj_val {
            if obj.as_any().downcast_ref::<Response>().is_some() {
                // 直接调用 Context 中的 Response 对象方法
                let result = self.context.response_mut().call_method(&method_lower, arg_values);

                // 检查是否需要设置退出标志（Response.End）
                if self.context.response().is_ended() {
                    self.context.should_exit = true;
                }

                return result;
            }
        }

        // 通用对象方法调用处理
        if let Value::Object(mut obj) = obj_val {
            let result = obj.call_method(&method_lower, arg_values);
            return result;
        }

        Ok(Value::Empty)
    }
}
