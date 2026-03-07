//! 声明语句执行
//!
//! 处理 Dim, Const, ReDim 和 Function/Sub 声明

use crate::ast::{Expr, Param, Stmt};
use crate::runtime::{Function, RuntimeError, Value};
use crate::utils::normalize_identifier;
use super::Interpreter;

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
            let mut size = 1;
            for dim_expr in sizes {
                let dim_val = self.eval_expr(dim_expr)?;
                let dim = match dim_val {
                    Value::Number(n) => n as usize,
                    _ => return Err(RuntimeError::Generic(format!(
                        "Array size must be a number, got {:?}",
                        dim_val
                    ))),
                };
                size *= dim.max(0); // 确保大小非负
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

            Value::Array(arr)
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
    pub fn eval_redim(&mut self, name: &str, sizes: &[Expr], preserve: bool) -> Result<Value, RuntimeError> {
        // 计算新数组大小
        let mut new_size = 1;
        for dim_expr in sizes {
            let dim_val = self.eval_expr(dim_expr)?;
            let dim = match dim_val {
                Value::Number(n) => n as usize,
                _ => return Err(RuntimeError::Generic(format!(
                    "Array size must be a number, got {:?}",
                    dim_val
                ))),
            };
            new_size *= dim.max(0);
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
        if let Some(Value::Array(old)) = old_arr {
            let copy_len = old.len().min(new_arr.len());
            for i in 0..copy_len {
                new_arr[i] = old[i].clone();
            }
        }

        // 设置新数组
        self.context.set_var(name.to_string(), Value::Array(new_arr));
        Ok(Value::Empty)
    }

    /// 注册函数(Sub 或 Function)
    pub fn register_function(&mut self, name: &str, params: &[Param], body: &[Stmt]) -> Result<Value, RuntimeError> {
        self.context.functions.insert(
            normalize_identifier(name),
            Function {
                name: name.to_string(),
                params: params.iter().map(|p| p.name.clone()).collect(),
                body: body.to_vec(),
            },
        );
        Ok(Value::Empty)
    }
}
