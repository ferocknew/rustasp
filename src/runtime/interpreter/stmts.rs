//! 语句执行模块
//!
//! 处理各种 VBScript 语句的执行逻辑

use crate::ast::{BinaryOp, Expr, IfBranch, Param, Stmt};
use crate::runtime::{BuiltinObject, Function, RuntimeError, Value, ValueCompare, ValueConversion};
use crate::utils::normalize_identifier;

use super::Interpreter;

/// 语句执行器
impl Interpreter {
    /// 执行语句（调度）
    pub fn eval_stmt(&mut self, stmt: &Stmt) -> Result<Value, RuntimeError> {
        match stmt {
            Stmt::Dim { name, init, is_array, sizes } => {
                self.eval_dim(name, init.as_ref(), *is_array, &sizes)
            }
            Stmt::Const { name, value } => self.eval_const(name, value),
            Stmt::Assignment { target, value } => self.eval_assignment(target, value),
            Stmt::Set { target, value } => self.eval_set(target, value),
            Stmt::If {
                branches,
                else_block,
            } => self.eval_if(branches, else_block),
            Stmt::For {
                var,
                start,
                end,
                step,
                body,
            } => self.eval_for(var, start, end, step.as_ref(), body),
            Stmt::While { cond, body } => self.eval_while(cond, body),
            Stmt::ForEach { var, collection, body } => self.eval_for_each(var, collection, body),
            Stmt::Select {
                expr,
                cases,
                else_block,
            } => self.eval_select(expr, cases, else_block),
            Stmt::ReDim { name, sizes, preserve } => self.eval_redim(name, sizes, *preserve),
            Stmt::Sub { name, params, body } | Stmt::Function { name, params, body } => {
                self.register_function(name, params, body)
            }
            Stmt::Call { name, args } => self.eval_call(name, args),
            Stmt::ExitFor => self.eval_exit_for(),
            Stmt::ExitFunction | Stmt::ExitSub => self.eval_exit(),
            Stmt::OptionExplicit => {
                // Option Explicit: 要求所有变量必须先声明
                // 当前实现暂时忽略，不强制检查
                Ok(Value::Empty)
            }
            Stmt::Expr(expr) => self.eval_expr(expr),
            _ => Err(RuntimeError::Generic(format!("Unimplemented: {:?}", stmt))),
        }
    }

    /// 执行 Dim 语句
    fn eval_dim(
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
    fn eval_const(&mut self, name: &str, value: &Expr) -> Result<Value, RuntimeError> {
        let val = self.eval_expr(value)?;
        self.context.define_var(name.to_string(), val);
        Ok(Value::Empty)
    }

    /// 执行赋值语句
    fn eval_assignment(&mut self, target: &Expr, value: &Expr) -> Result<Value, RuntimeError> {
        let val = self.eval_expr(value)?;

        match target {
            Expr::Variable(name) => {
                self.context.set_var(name.clone(), val);
                Ok(Value::Empty)
            }
            Expr::Index { object, index } => self.eval_index_assignment(object, index, val),
            Expr::Property { object, property } => self.eval_property_assignment(object, property, val),
            _ => Err(RuntimeError::InvalidAssignment),
        }
    }

    /// 执行 Set 语句
    fn eval_set(&mut self, target: &Expr, value: &Expr) -> Result<Value, RuntimeError> {
        self.eval_assignment(target, value)
    }

    /// 执行 If 语句
    fn eval_if(
        &mut self,
        branches: &[IfBranch],
        else_block: &Option<Vec<Stmt>>,
    ) -> Result<Value, RuntimeError> {
        for branch in branches {
            let cond = self.eval_expr(&branch.cond)?;
            if cond.is_truthy() {
                return self.exec_block(&branch.body);
            }
        }
        if let Some(else_stmts) = else_block {
            self.exec_block(else_stmts)?;
        }
        Ok(Value::Empty)
    }

    /// 执行 For 循环
    fn eval_for(
        &mut self,
        var: &str,
        start: &Expr,
        end: &Expr,
        step: Option<&Expr>,
        body: &[Stmt],
    ) -> Result<Value, RuntimeError> {
        let start_val = self.eval_expr(start)?.to_number();
        let end_val = self.eval_expr(end)?.to_number();
        let step_val = step
            .map(|e| self.eval_expr(e).map(|v: Value| v.to_number()))
            .transpose()?
            .unwrap_or(1.0);

        let mut i = start_val;
        let condition = if step_val > 0.0 {
            move |i: f64, end: f64| i <= end
        } else {
            move |i: f64, end: f64| i >= end
        };

        while condition(i, end_val) {
            self.context.define_var(var.to_string(), Value::Number(i));
            self.exec_block(body)?;
            i += step_val;
        }

        Ok(Value::Empty)
    }

    /// 执行 While 循环
    fn eval_while(&mut self, cond: &Expr, body: &[Stmt]) -> Result<Value, RuntimeError> {
        while self.eval_expr(cond)?.is_truthy() {
            self.exec_block(body)?;
        }
        Ok(Value::Empty)
    }

    /// 执行 Select Case 语句
    fn eval_select(
        &mut self,
        expr: &Expr,
        cases: &[crate::ast::CaseClause],
        else_block: &Option<Vec<Stmt>>,
    ) -> Result<Value, RuntimeError> {
        let select_value = self.eval_expr(expr)?;

        for case in cases {
            if let Some(values) = &case.values {
                for value_expr in values {
                    let case_value = self.eval_expr(value_expr)?;
                    let result = select_value.compare(BinaryOp::Eq, &case_value);
                    if let Value::Boolean(true) = result {
                        return self.exec_block(&case.body);
                    }
                }
            }
        }

        if let Some(else_stmts) = else_block {
            self.exec_block(else_stmts)?;
        }

        Ok(Value::Empty)
    }

    /// 执行 Call 语句
    fn eval_call(&mut self, name: &str, args: &[Expr]) -> Result<Value, RuntimeError> {
        let arg_values: Result<Vec<Value>, _> = args.iter().map(|e| self.eval_expr(e)).collect();
        let arg_values = arg_values?;

        let name_lower = normalize_identifier(name);
        if let Some(func) = self.context.functions.get(&name_lower).cloned() {
            self.context.push_scope();

            for (i, param_name) in func.params.iter().enumerate() {
                let value = if i < arg_values.len() {
                    arg_values[i].clone()
                } else {
                    Value::Empty
                };
                self.context.define_var(param_name.clone(), value);
            }

            let result = self.exec_block(&func.body);
            self.context.pop_scope();
            result?;
        }
        Ok(Value::Empty)
    }

    /// Exit For
    fn eval_exit_for(&mut self) -> Result<Value, RuntimeError> {
        // TODO: 实现循环退出
        Ok(Value::Empty)
    }

    /// Exit Function/Sub
    fn eval_exit(&mut self) -> Result<Value, RuntimeError> {
        self.context.should_exit = true;
        Ok(Value::Empty)
    }

    // ==================== 内建对象属性设置 ====================

    /// 设置 Response 对象的属性
    fn builtin_response_set_property(&mut self, property: &str, value: Value) -> Result<Value, RuntimeError> {
        let response = self.context.response_mut();
        match property.to_uppercase().as_str() {
            "BUFFER" | "CONTENTTYPE" | "CACHECONTROL" | "EXPIRES" | "EXPIRESABSOLUTE" | "PICS" | "ISCLIENTCONNECTED" => {
                let prop_lower = property.to_lowercase();
                response.set_property(&prop_lower, value)?;
                Ok(Value::Empty)
            }
            "CHARSET" | "STATUS" => Ok(Value::Empty),
            _ => Err(RuntimeError::PropertyNotFound(format!("Response.{}", property))),
        }
    }

    /// 设置 Server 对象的属性
    fn builtin_server_set_property(&mut self, property: &str, _value: Value) -> Result<Value, RuntimeError> {
        match property.to_uppercase().as_str() {
            "SCRIPTTIMEOUT" => Ok(Value::Empty),
            _ => Err(RuntimeError::PropertyNotFound(format!("Server.{}", property))),
        }
    }

    /// 设置 Session 对象的属性
    fn builtin_session_set_property(&mut self, property: &str, value: Value) -> Result<Value, RuntimeError> {
        match property.to_uppercase().as_str() {
            "TIMEOUT" | "CODEPAGE" | "LCID" => Ok(Value::Empty),
            _ => {
                if let Some(Value::Object(mut map)) = self.context.get_var("Session").cloned() {
                    map.insert(property.to_lowercase(), value);
                    self.context.set_var("Session".to_string(), Value::Object(map));
                    Ok(Value::Empty)
                } else {
                    Err(RuntimeError::Generic("Session object not found".to_string()))
                }
            }
        }
    }

    /// 执行 ReDim 语句
    fn eval_redim(&mut self, name: &str, sizes: &[Expr], preserve: bool) -> Result<Value, RuntimeError> {
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

    /// 执行 For Each 循环
    fn eval_for_each(&mut self, var: &str, collection: &Expr, body: &[Stmt]) -> Result<Value, RuntimeError> {
        let collection_val = self.eval_expr(collection)?;

        let elements = match collection_val {
            Value::Array(arr) => arr,
            Value::Object(obj) => obj.values().cloned().collect::<Vec<_>>(),
            Value::String(s) => s.chars().map(|c| Value::String(c.to_string())).collect(),
            _ => {
                return Err(RuntimeError::Generic(format!(
                    "For Each requires an array, object, or string, got {:?}",
                    collection_val
                )))
            }
        };

        for element in elements {
            self.context.define_var(var.to_string(), element);
            self.exec_block(body)?;
        }

        Ok(Value::Empty)
    }

    // ==================== 辅助方法 ====================

    /// 执行语句块
    fn exec_block(&mut self, stmts: &[Stmt]) -> Result<Value, RuntimeError> {
        for stmt in stmts {
            self.eval_stmt(stmt)?;
        }
        Ok(Value::Empty)
    }

    /// 注册函数(Sub 或 Function)
    fn register_function(&mut self, name: &str, params: &[Param], body: &[Stmt]) -> Result<Value, RuntimeError> {
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

    /// 执行索引赋值（如 arr(i) = value 或 Session("key") = value）
    fn eval_index_assignment(&mut self, object: &Expr, index: &Expr, val: Value) -> Result<Value, RuntimeError> {
        let idx = self.eval_expr(index)?;

        match object {
            Expr::Variable(name) => {
                if name.to_lowercase() == "session" {
                    if let Value::String(key) = idx {
                        return self.builtin_session_set_property(&key, val);
                    }
                }

                if let Value::Number(i) = idx {
                    let i = i as usize;
                    match self.context.get_var(name).cloned() {
                        Some(Value::Array(mut arr)) => {
                            if i >= arr.len() {
                                arr.resize(i + 1, Value::Empty);
                            }
                            arr[i] = val;
                            self.context.set_var(name.clone(), Value::Array(arr));
                            Ok(Value::Empty)
                        }
                        Some(Value::Empty) => {
                            let mut arr = vec![Value::Empty; i + 1];
                            arr[i] = val;
                            self.context.set_var(name.clone(), Value::Array(arr));
                            Ok(Value::Empty)
                        }
                        _ => {
                            let mut arr = vec![Value::Empty; i + 1];
                            arr[i] = val;
                            self.context.set_var(name.clone(), Value::Array(arr));
                            Ok(Value::Empty)
                        }
                    }
                } else {
                    Err(RuntimeError::Generic(format!("Array index must be a number, got {:?}", idx)))
                }
            }
            _ => Err(RuntimeError::InvalidAssignment),
        }
    }

    /// 执行属性赋值（如 Response.Buffer = value）
    fn eval_property_assignment(&mut self, object: &Expr, property: &str, val: Value) -> Result<Value, RuntimeError> {
        match object {
            Expr::Variable(obj_name) => {
                match obj_name.to_lowercase().as_str() {
                    "response" => self.builtin_response_set_property(property, val),
                    "request" => Err(RuntimeError::PropertyNotFound(format!("Request.{}", property))),
                    "server" => self.builtin_server_set_property(property, val),
                    "session" => self.builtin_session_set_property(property, val),
                    _ => Err(RuntimeError::PropertyNotFound(format!("{}.{}", obj_name, property))),
                }
            }
            _ => Err(RuntimeError::InvalidAssignment),
        }
    }
}
