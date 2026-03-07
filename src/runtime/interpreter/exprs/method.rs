//! 方法调用表达式求值

use crate::ast::Expr;
use crate::runtime::{BuiltinObject, RuntimeError, Value};

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

        // 特殊处理：Response 对象（存储在 context 中，不在变量表中）
        if let Expr::Variable(name) = object {
            if crate::utils::normalize_identifier(name) == "response" {
                return self.eval_response_method(&method_lower, arg_values);
            }
        }

        // 统一 trait dispatch：eval object 并调用方法
        let obj_val = self.eval_expr(object)?;
        
        if let Value::Object(mut obj) = obj_val {
            return obj.call_method(&method_lower, arg_values);
        }

        Ok(Value::Empty)
    }

    /// 处理 Response 对象的方法调用
    fn eval_response_method(&mut self, method: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        let response = self.context.response_mut();
        let result = response.call_method(method, args);
        
        // 检查 Response.End 标志
        if response.is_ended() {
            self.context.should_exit = true;
        }
        
        result
    }
}
