//! 索引访问表达式求值

use crate::ast::Expr;
use crate::runtime::{RuntimeError, Value};

use super::super::Interpreter;

impl Interpreter {
    /// 处理索引表达式（支持多维索引）
    pub fn eval_index(&mut self, object: &Expr, indices: &[Expr]) -> Result<Value, RuntimeError> {
        // 求值所有索引
        let index_vals: Result<Vec<Value>, RuntimeError> =
            indices.iter().map(|idx| self.eval_expr(idx)).collect();
        let index_vals = index_vals?;

        match object {
            // 处理变量索引访问
            Expr::Variable(name) => {
                // 单个索引：检查是否是内置函数调用
                if indices.len() == 1 {
                    if let Some(result) = self.call_builtin(name, &[index_vals[0].clone()]) {
                        return result;
                    }
                }

                // 检查变量是否是数组或对象
                if let Some(value) = self.context.get_var(name) {
                    return self.eval_value_index(value.clone(), &index_vals);
                }
                Err(RuntimeError::InvalidIndex)
            }
            // 处理嵌套表达式
            _ => {
                let obj_val = self.eval_expr(object)?;
                self.eval_value_index(obj_val, &index_vals)
            }
        }
    }

    /// 对值执行索引访问
    fn eval_value_index(&self, value: Value, indices: &[Value]) -> Result<Value, RuntimeError> {
        match value {
            Value::Array(ref arr) => {
                // 将索引转换为 usize
                let idx: Result<Vec<usize>, RuntimeError> = indices
                    .iter()
                    .map(|v| match v {
                        Value::Number(n) => Ok(*n as usize),
                        _ => Err(RuntimeError::Generic(format!(
                            "Array index must be a number, got {:?}",
                            v
                        ))),
                    })
                    .collect();

                let idx = idx?;

                // 使用 flat_index 计算扁平索引
                let locked_arr = arr
                    .lock()
                    .map_err(|_| RuntimeError::Generic("Failed to lock array".to_string()))?;

                match locked_arr.flat_index(&idx) {
                    Some(flat_idx) => Ok(locked_arr.data[flat_idx].clone()),
                    None => Ok(Value::Empty),
                }
            }
            Value::Object(ref obj) => {
                // 对象索引：只支持单个索引（如 Session("key")）
                if indices.len() == 1 {
                    obj.lock()
                        .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?
                        .index(&indices[0])
                } else {
                    Err(RuntimeError::Generic(
                        "Object index only supports single dimension".to_string(),
                    ))
                }
            }
            _ => Err(RuntimeError::InvalidIndex),
        }
    }
}
