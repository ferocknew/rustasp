//! 索引访问表达式求值

use crate::ast::Expr;
use crate::runtime::{RuntimeError, Value, ValueIndex};

use super::super::Interpreter;

impl Interpreter {
    /// 处理索引表达式
    pub fn eval_index(&mut self, object: &Expr, index: &Expr) -> Result<Value, RuntimeError> {
        let index_val = self.eval_expr(index)?;

        match object {
            // 处理变量索引访问
            Expr::Variable(name) => {
                // 检查是否是内置函数调用
                if let Some(result) = self.call_builtin(name, &[index_val.clone()]) {
                    return result;
                }

                // 检查变量是否是数组或对象
                if let Some(value) = self.context.get_var(name).cloned() {
                    return value.index(&index_val);
                }
                Err(RuntimeError::InvalidIndex)
            }
            // 处理嵌套表达式
            _ => {
                let obj_val = self.eval_expr(object)?;
                obj_val.index(&index_val)
            }
        }
    }
}
