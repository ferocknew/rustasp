//! 控制流语句执行模块
//!
//! 处理 If、For、While、ForEach、Select Case 等控制流语句

use crate::ast::{BinaryOp, CaseClause, Expr, IfBranch, Stmt};
use crate::runtime::{ControlFlow, RuntimeError, Value, ValueCompare, ValueConversion};

use super::Interpreter;

/// 控制流语句执行器
impl Interpreter {
    /// 执行 If 语句
    pub fn eval_if(
        &mut self,
        branches: &[IfBranch],
        else_block: &Option<Vec<crate::ast::Stmt>>,
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
    pub fn eval_for(
        &mut self,
        var: &str,
        start: &Expr,
        end: &Expr,
        step: Option<&Expr>,
        body: &[crate::ast::Stmt],
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
            match self.exec_block(body) {
                Ok(_) => {}
                Err(RuntimeError::ControlFlow(ControlFlow::ExitFor)) => break,
                Err(e) => return Err(e),
            }
            i += step_val;
        }

        Ok(Value::Empty)
    }

    /// 执行 While 循环
    pub fn eval_while(
        &mut self,
        cond: &Expr,
        body: &[crate::ast::Stmt],
    ) -> Result<Value, RuntimeError> {
        while self.eval_expr(cond)?.is_truthy() {
            match self.exec_block(body) {
                Ok(_) => {}
                Err(RuntimeError::ControlFlow(ControlFlow::ExitDo)) => break,
                Err(e) => return Err(e),
            }
        }
        Ok(Value::Empty)
    }

    /// 执行 Do While 循环
    pub fn eval_do_while(
        &mut self,
        cond: &Expr,
        body: &[Stmt],
    ) -> Result<Value, RuntimeError> {
        while self.eval_expr(cond)?.is_truthy() {
            match self.exec_block(body) {
                Ok(_) => {}
                Err(RuntimeError::ControlFlow(ControlFlow::ExitDo)) => break,
                Err(e) => return Err(e),
            }
        }
        Ok(Value::Empty)
    }

    /// 执行 Do Until 循环
    pub fn eval_do_until(
        &mut self,
        cond: &Expr,
        body: &[Stmt],
    ) -> Result<Value, RuntimeError> {
        while !self.eval_expr(cond)?.is_truthy() {
            match self.exec_block(body) {
                Ok(_) => {}
                Err(RuntimeError::ControlFlow(ControlFlow::ExitDo)) => break,
                Err(e) => return Err(e),
            }
        }
        Ok(Value::Empty)
    }

    /// 执行 Do...Loop While
    pub fn eval_do_loop_while(
        &mut self,
        body: &[Stmt],
        cond: &Expr,
    ) -> Result<Value, RuntimeError> {
        loop {
            match self.exec_block(body) {
                Ok(_) => {}
                Err(RuntimeError::ControlFlow(ControlFlow::ExitDo)) => break,
                Err(e) => return Err(e),
            }
            if !self.eval_expr(cond)?.is_truthy() {
                break;
            }
        }
        Ok(Value::Empty)
    }

    /// 执行 Do...Loop Until
    pub fn eval_do_loop_until(
        &mut self,
        body: &[Stmt],
        cond: &Expr,
    ) -> Result<Value, RuntimeError> {
        loop {
            match self.exec_block(body) {
                Ok(_) => {}
                Err(RuntimeError::ControlFlow(ControlFlow::ExitDo)) => break,
                Err(e) => return Err(e),
            }
            if self.eval_expr(cond)?.is_truthy() {
                break;
            }
        }
        Ok(Value::Empty)
    }

    /// 执行 Do...Loop (无限循环)
    pub fn eval_do_loop(&mut self, body: &[Stmt]) -> Result<Value, RuntimeError> {
        loop {
            match self.exec_block(body) {
                Ok(_) => {}
                Err(RuntimeError::ControlFlow(ControlFlow::ExitDo)) => break,
                Err(e) => return Err(e),
            }
        }
        Ok(Value::Empty)
    }

    /// 执行 For Each 循环
    pub fn eval_for_each(
        &mut self,
        var: &str,
        collection: &Expr,
        body: &[crate::ast::Stmt],
    ) -> Result<Value, RuntimeError> {
        let collection_val = self.eval_expr(collection)?;

        let elements = match collection_val {
            Value::Array(ref arr) => {
                let locked_arr = arr.lock()
                    .map_err(|_| RuntimeError::Generic("Failed to lock array".to_string()))?;
                locked_arr.data.clone()
            }
            Value::Object(ref obj) => {
                // 尝试作为字典处理
                use crate::runtime::objects::Dictionary;
                let locked_obj = obj.lock()
                    .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?;

                if let Some(dict) = locked_obj.as_any().downcast_ref::<Dictionary>() {
                    dict.values().to_vec()
                } else {
                    // 对于其他对象，尝试调用 items 方法
                    drop(locked_obj);
                    match obj.lock()
                        .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?
                        .call_method("items", vec![]) {
                        Ok(Value::Array(ref arr)) => {
                            let locked_arr = arr.lock()
                                .map_err(|_| RuntimeError::Generic("Failed to lock array".to_string()))?;
                            locked_arr.data.clone()
                        }
                        _ => {
                            return Err(RuntimeError::Generic(
                                "For Each requires an iterable object".to_string(),
                            ))
                        }
                    }
                }
            }
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
            match self.exec_block(body) {
                Ok(_) => {}
                Err(RuntimeError::ControlFlow(ControlFlow::ExitFor)) => break,
                Err(e) => return Err(e),
            }
        }

        Ok(Value::Empty)
    }

    /// 执行 Select Case 语句
    pub fn eval_select(
        &mut self,
        expr: &Expr,
        cases: &[CaseClause],
        else_block: &Option<Vec<crate::ast::Stmt>>,
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
    pub fn eval_call(
        &mut self,
        name: &str,
        args: &[Expr],
    ) -> Result<Value, RuntimeError> {
        let name_lower = crate::utils::normalize_identifier(name);
        if let Some(func) = self.context.functions.get(&name_lower).cloned() {
            // 记录 ByRef 参数映射: (参数索引 -> 原始变量名)
            let mut byref_mapping: Vec<(String, String)> = Vec::new();
            
            // 计算参数值
            let mut arg_values = Vec::new();
            for (i, arg) in args.iter().enumerate() {
                // 检查是否是 ByRef 参数且参数为变量
                let is_byref_param = i < func.params.len() && func.params[i].is_byref;
                if is_byref_param {
                    if let Expr::Variable(var_name) = arg {
                        // ByRef 参数，记录映射
                        let param_name = func.params[i].name.clone();
                        byref_mapping.push((param_name, var_name.clone()));
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
            self.context.define_var(func.name.clone(), Value::Empty);

            // 执行函数体，处理 Exit Sub/Function
            for stmt in &func.body {
                match self.eval_stmt(stmt) {
                    Ok(_) => {}
                    Err(RuntimeError::ControlFlow(ControlFlow::ExitFunction)) |
                    Err(RuntimeError::ControlFlow(ControlFlow::ExitSub)) => {
                        // Exit Function/Sub - 正常退出
                        break;
                    }
                    Err(e) => {
                        self.context.pop_scope();
                        return Err(e);
                    }
                }
            }

            // 在 pop_scope 之前保存 ByRef 参数的值
            let byref_values: Vec<(String, Value)> = byref_mapping.iter()
                .filter_map(|(param_name, _)| {
                    self.context.get_var(param_name).cloned()
                        .map(|v| (param_name.clone(), v))
                })
                .collect();

            self.context.pop_scope();

            // 在 pop_scope 之后，将 ByRef 参数的值写回外部变量
            for (param_name, original_var_name) in &byref_mapping {
                if let Some((_, value)) = byref_values.iter().find(|(pn, _)| pn == param_name) {
                    self.context.set_var(original_var_name.clone(), value.clone());
                }
            }
        }
        Ok(Value::Empty)
    }

    /// Exit For
    pub fn eval_exit_for(&mut self) -> Result<Value, RuntimeError> {
        Err(RuntimeError::ControlFlow(ControlFlow::ExitFor))
    }

    /// Exit Do
    pub fn eval_exit_do(&mut self) -> Result<Value, RuntimeError> {
        Err(RuntimeError::ControlFlow(ControlFlow::ExitDo))
    }

    /// Exit Function
    pub fn eval_exit_function(&mut self) -> Result<Value, RuntimeError> {
        Err(RuntimeError::ControlFlow(ControlFlow::ExitFunction))
    }

    /// Exit Sub
    pub fn eval_exit_sub(&mut self) -> Result<Value, RuntimeError> {
        Err(RuntimeError::ControlFlow(ControlFlow::ExitSub))
    }

    /// Exit Property
    pub fn eval_exit_property(&mut self) -> Result<Value, RuntimeError> {
        Err(RuntimeError::ControlFlow(ControlFlow::ExitProperty))
    }
}
