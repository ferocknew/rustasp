//! 方法调用表达式求值

use crate::ast::Expr;
use crate::runtime::{RuntimeError, Value};
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

        // 特殊处理：如果是 Response 对象，直接操作变量表中的对象
        if let Value::Object(ref obj) = obj_val {
            if obj.as_any().downcast_ref::<Response>().is_some() {
                // 直接从变量表获取 Response 对象的可变引用
                if let Some(Value::Object(response_obj)) = self.context.get_var_mut("Response") {
                    let result = response_obj.call_method(&method_lower, arg_values);

                    // 检查是否需要设置退出标志（Response.End）
                    if let Some(resp) = response_obj.as_any().downcast_ref::<Response>() {
                        if resp.is_ended() {
                            self.context.should_exit = true;
                        }
                    }

                    return result;
                }
            }
        }

        // 通用对象方法调用处理
        // 需要将修改后的对象写回变量表
        if let Expr::Variable(name) = object {
            if let Some(Value::Object(ref mut obj)) = self.context.get_var_mut(name) {
                let result = obj.call_method(&method_lower, arg_values);
                return result;
            }
        }

        // 对于其他情况（如链式调用），使用临时对象
        if let Value::Object(mut obj) = obj_val {
            let result = obj.call_method(&method_lower, arg_values);
            return result;
        }

        Ok(Value::Empty)
    }
}
