//! 声明语句执行模块
//!
//! 处理 Dim、Const、ReDim 等变量声明语句

use crate::ast::Expr;
use crate::runtime::{RuntimeError, Value, VbsArray};
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
            let mut dims = vec![];
            for dim_expr in sizes {
                let dim_val = self.eval_expr(dim_expr)?;
                let dim = match dim_val {
                    Value::Number(n) => n as usize + 1,  // +1 因为 0-based
                    _ => {
                        return Err(RuntimeError::Generic(format!(
                            "Array size must be a number, got {:?}",
                            dim_val
                        )))
                    }
                };
                dims.push(dim);
            }

            // 创建 VbsArray
            let mut vbs_arr = VbsArray::new(dims);

            // 如果有初始化值，填充第一个元素
            if let Some(init_expr) = init {
                let init_val = self.eval_expr(init_expr)?;
                if !vbs_arr.data.is_empty() {
                    vbs_arr.data[0] = init_val;
                }
            }

            Value::Array(Arc::new(Mutex::new(vbs_arr)))
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
        // 计算新数组维度
        // VBScript: ReDim arr(n) 创建索引 0 到 n 的数组，大小为 n+1
        let mut new_dims = vec![];
        for dim_expr in sizes {
            let dim_val = self.eval_expr(dim_expr)?;
            let dim = match dim_val {
                Value::Number(n) => n as usize + 1,  // +1 因为 0-based
                _ => {
                    return Err(RuntimeError::Generic(format!(
                        "Array size must be a number, got {:?}",
                        dim_val
                    )))
                }
            };
            new_dims.push(dim);
        }

        // 检查变量是否存在
        if let Some(Value::Array(ref arr_ref)) = self.context.get_var(name) {
            // 使用 VbsArray::redim 方法
            let mut arr = arr_ref.lock()
                .map_err(|_| RuntimeError::Generic("Failed to lock array".to_string()))?;
            arr.redim(new_dims, preserve);
            Ok(Value::Empty)
        } else {
            // 数组不存在，创建新数组
            let vbs_arr = VbsArray::new(new_dims);
            self.context.set_var(name.to_string(), Value::Array(Arc::new(Mutex::new(vbs_arr))));
            Ok(Value::Empty)
        }
    }
}
