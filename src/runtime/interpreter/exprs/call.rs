//! 函数调用表达式求值

use crate::ast::Expr;
use crate::runtime::{ControlFlow, RuntimeError, Value, ValueIndex};
use crate::runtime::builtins::BuiltinExecutor;

use super::super::Interpreter;

impl Interpreter {
    /// 执行函数调用表达式
    pub fn eval_call_expr(&mut self, name: &str, args: &[Expr]) -> Result<Value, RuntimeError> {
        // 优化：先检查是否是内置函数，避免不必要的参数 eval
        if let Some(token) = self.builtin_registry.lookup(name) {
            // 内置函数：先 eval 参数再调用
            let arg_values: Result<Vec<Value>, _> = args.iter().map(|e| self.eval_expr(e)).collect();
            let arg_values = arg_values?;
            return BuiltinExecutor::execute(token, &arg_values);
        }

        // 用户函数调用，传递原始表达式以支持 ByRef
        self.eval_user_function_call(name, args)
    }

    /// 用户函数调用
    fn eval_user_function_call(&mut self, name: &str, args: &[Expr]) -> Result<Value, RuntimeError> {
        // 检查是否是用户函数
        let func = self.context.get_function(name).cloned();
        
        if let Some(func) = func {
            // 记录 ByRef 参数映射: (参数索引, 原始变量名) - 优化：使用索引而非参数名
            let mut byref_indices: Vec<(usize, String)> = Vec::new();
            
            // 计算参数值（优化：只 eval 一次）
            let mut arg_values = Vec::new();
            for (i, arg) in args.iter().enumerate() {
                // 检查是否是 ByRef 参数且参数为变量
                let is_byref_param = i < func.params.len() && func.params[i].is_byref;
                if is_byref_param {
                    if let Expr::Variable(var_name) = arg {
                        // ByRef 参数，记录索引和变量名
                        byref_indices.push((i, var_name.clone()));
                        // 使用当前变量值
                        let value = self.context.get_var(var_name)
                            .cloned()
                            .unwrap_or(Value::Empty);
                        arg_values.push(value);
                    } else {
                        // 不是变量表达式，按值传递
                        arg_values.push(self.eval_expr(arg)?);
                    }
                } else {
                    // ByVal 参数，正常计算
                    arg_values.push(self.eval_expr(arg)?);
                }
            }

            self.context.push_scope();

            // 绑定参数
            for (i, param) in func.params.iter().enumerate() {
                let value = if i < arg_values.len() {
                    arg_values[i].clone()
                } else {
                    Value::Empty
                };
                self.context.define_var(param.name.clone(), value);
            }

            // 初始化函数名变量为 Empty（用于返回值）
            let func_name_lower = crate::utils::normalize_identifier(&func.name);
            self.context.define_var(func.name.clone(), Value::Empty);

            // 执行函数体，处理 Exit Function/Sub
            for stmt in &func.body {
                match self.eval_stmt(stmt) {
                    Ok(_) => {}
                    Err(RuntimeError::ControlFlow(ControlFlow::ExitFunction)) |
                    Err(RuntimeError::ControlFlow(ControlFlow::ExitSub)) => {
                        // Exit Function/Sub - 正常退出，读取返回值
                        break;
                    }
                    Err(e) => {
                        self.context.pop_scope();
                        return Err(e);
                    }
                }
            }

            // 获取返回值（Function 名称变量的值）
            let result = self.context.get_var(&func.name)
                .or_else(|| self.context.get_var(&func_name_lower))
                .cloned()
                .unwrap_or(Value::Empty);

            // 在 pop_scope 之前保存 ByRef 参数的值（优化：使用索引直接访问）
            let byref_values: Vec<Value> = byref_indices.iter()
                .map(|(idx, _)| {
                    let param_name = &func.params[*idx].name;
                    self.context.get_var(param_name).cloned().unwrap_or(Value::Empty)
                })
                .collect();

            self.context.pop_scope();

            // 在 pop_scope 之后，将 ByRef 参数的值写回外部变量（优化：O(1) 索引访问）
            for (i, (_, original_var_name)) in byref_indices.iter().enumerate() {
                self.context.set_var(original_var_name.clone(), byref_values[i].clone());
            }

            return Ok(result);
        }

        // 回退到数组索引访问（需要 eval args）
        if args.len() == 1 {
            let index = self.eval_expr(&args[0])?;
            // 处理数组索引访问：arr(0)
            if let Some(value) = self.context.get_var(name).cloned() {
                return value.index(&index);
            }
        }

        Err(RuntimeError::UndefinedVariable(format!("Function '{}' or array index", name)))
    }

    /// 尝试调用内置函数
    pub fn call_builtin(&mut self, name: &str, args: &[Value]) -> Option<Result<Value, RuntimeError>> {
        self.builtin_registry.lookup(name).map(|token| BuiltinExecutor::execute(token, args))
    }
}
