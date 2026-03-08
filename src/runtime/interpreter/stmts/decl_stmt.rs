//! 声明语句执行模块
//!
//! 处理 Dim、Const、ReDim 等变量声明语句

use crate::ast::Expr;
use crate::runtime::{RuntimeError, Value};
use std::sync::{Arc, Mutex};

use super::Interpreter;

/// 声明语句执行器
impl Interpreter {
    /// 执行 Dim 语句
    pub fn eval_dim(
        &mut self,
        name: &str,
        init: Option<&Expr>,
        is_array: bool,
        sizes: &[Expr],
    ) -> Result<Value, RuntimeError> {
        let value = if is_array {
            // 创建数组
            // VBScript: ReDim arr(n) 创建索引 0 到 n 的数组，大小为 n+1
            let mut size = 0;
            for dim_expr in sizes {
                let dim_val = self.eval_expr(dim_expr)?;
                let dim = match dim_val {
                    Value::Number(n) => n as usize,
                    _ => {
                        return Err(RuntimeError::Generic(format!(
                            "Array size must be a number, got {:?}",
                            dim_val
                        )))
                    }
                };
                // ReDim arr(n) 表示最大索引是 n，所以大小是 n+1
                if size == 0 {
                    size = dim + 1;
                } else {
                    size *= dim + 1;
                }
            }

            // 创建指定大小的空数组
            let mut arr = vec![Value::Empty; size];
            // 如果有初始化值，填充第一个元素
            if let Some(init_expr) = init {
                let init_val = self.eval_expr(init_expr)?;
                if !arr.is_empty() {
                    arr[0] = init_val;
                }
            }

            Value::Array(Arc::new(Mutex::new(arr)))
        } else if let Some(expr) = init {
            self.eval_expr(expr)?
        } else {
            Value::Empty
        };

        self.context.define_var(name.to_string(), value);
        Ok(Value::Empty)
    }

    /// 执行 Const 语句
    pub fn eval_const(&mut self, name: &str, value: &Expr) -> Result<Value, RuntimeError> {
        let val = self.eval_expr(value)?;
        self.context.define_var(name.to_string(), val);
        Ok(Value::Empty)
    }

    /// 执行 ReDim 语句
    pub fn eval_redim(
        &mut self,
        name: &str,
        sizes: &[Expr],
        preserve: bool,
    ) -> Result<Value, RuntimeError> {
        // 计算新数组大小
        // VBScript: ReDim arr(n) 创建索引 0 到 n 的数组，大小为 n+1
        let mut new_size = 0;
        for dim_expr in sizes {
            let dim_val = self.eval_expr(dim_expr)?;
            let dim = match dim_val {
                Value::Number(n) => n as usize,
                _ => {
                    return Err(RuntimeError::Generic(format!(
                        "Array size must be a number, got {:?}",
                        dim_val
                    )))
                }
            };
            // ReDim arr(n) 表示最大索引是 n，所以大小是 n+1
            if new_size == 0 {
                new_size = dim + 1;
            } else {
                new_size *= dim + 1;
            }
        }

        // 获取旧数组（如果需要 preserve）
        let old_arr = if preserve {
            self.context.get_var(name).cloned()
        } else {
            None
        };

        // 创建新数组
        let mut new_arr = vec![Value::Empty; new_size];

        // 如果需要 preserve，复制旧数据
        if let Some(Value::Array(ref old_arr)) = old_arr {
            let locked_old = old_arr.lock().unwrap();
            let copy_len = locked_old.len().min(new_arr.len());
            for i in 0..copy_len {
                new_arr[i] = locked_old[i].clone();
            }
        }

        // 设置新数组
        self.context.set_var(name.to_string(), Value::Array(Arc::new(Mutex::new(new_arr))));
        Ok(Value::Empty)
    }
}
