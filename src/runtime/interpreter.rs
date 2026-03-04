//! 解释器 - 调度语句执行

use super::value::{ValueCompare, ValueConversion, ValueOps};
use super::{Context, Function, RuntimeError, Value};
use crate::ast::{BinaryOp, Expr, Program, Stmt, UnaryOp};

/// 解释器
pub struct Interpreter {
    context: Context,
}

impl Interpreter {
    /// 创建新解释器
    pub fn new() -> Self {
        Interpreter {
            context: Context::new(),
        }
    }

    /// 获取上下文
    pub fn context(&self) -> &Context {
        &self.context
    }

    /// 获取可变上下文
    pub fn context_mut(&mut self) -> &mut Context {
        &mut self.context
    }

    /// 执行程序
    pub fn execute(&mut self, program: &Program) -> Result<Value, RuntimeError> {
        let mut result = Value::Empty;
        for stmt in &program.statements {
            result = self.eval_stmt(stmt)?;
            if self.context.should_exit {
                break;
            }
        }
        Ok(result)
    }

    /// 执行语句（调度）
    fn eval_stmt(&mut self, stmt: &Stmt) -> Result<Value, RuntimeError> {
        match stmt {
            Stmt::Dim { name, init, .. } => self.eval_dim(name, init.as_ref()),
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
            Stmt::Select {
                expr,
                cases,
                else_block,
            } => self.eval_select(expr, cases, else_block),
            Stmt::Sub { name, params, body } => self.eval_sub(name, params, body),
            Stmt::Function { name, params, body } => self.eval_function(name, params, body),
            Stmt::Call { name, args } => self.eval_call(name, args),
            Stmt::ExitFor => self.eval_exit_for(),
            Stmt::ExitFunction => self.eval_exit_function(),
            Stmt::ExitSub => self.eval_exit_sub(),
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
    fn eval_dim(&mut self, name: &str, init: Option<&Expr>) -> Result<Value, RuntimeError> {
        let value = if let Some(expr) = init {
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
            Expr::Index { object, index } => {
                let idx = self.eval_expr(index)?;
                match object.as_ref() {
                    Expr::Variable(name) => {
                        if let Some(Value::Array(arr)) = self.context.get_var(name).cloned() {
                            let mut arr = arr;
                            if let Value::Number(i) = idx {
                                let i = i as usize;
                                if i < arr.len() {
                                    arr[i] = val;
                                    self.context.set_var(name.clone(), Value::Array(arr));
                                    return Ok(Value::Empty);
                                }
                            }
                        }
                        Err(RuntimeError::InvalidAssignment)
                    }
                    _ => Err(RuntimeError::InvalidAssignment),
                }
            }
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
        branches: &[crate::ast::IfBranch],
        else_block: &Option<Vec<Stmt>>,
    ) -> Result<Value, RuntimeError> {
        for branch in branches {
            let cond = self.eval_expr(&branch.cond)?;
            if cond.is_truthy() {
                for stmt in &branch.body {
                    self.eval_stmt(stmt)?;
                }
                return Ok(Value::Empty);
            }
        }
        if let Some(else_stmts) = else_block {
            for stmt in else_stmts {
                self.eval_stmt(stmt)?;
            }
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
            .map(|e| self.eval_expr(e).map(|v| v.to_number()))
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
            for stmt in body {
                self.eval_stmt(stmt)?;
            }
            i += step_val;
        }

        Ok(Value::Empty)
    }

    /// 执行 While 循环
    fn eval_while(&mut self, cond: &Expr, body: &[Stmt]) -> Result<Value, RuntimeError> {
        while self.eval_expr(cond)?.is_truthy() {
            for stmt in body {
                self.eval_stmt(stmt)?;
            }
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
        // 计算表达式的值
        let select_value = self.eval_expr(expr)?;

        // 遍历所有 Case 分支
        for case in cases {
            if let Some(values) = &case.values {
                // 检查是否有匹配的值
                for value_expr in values {
                    let case_value = self.eval_expr(value_expr)?;
                    // 使用 compare 方法进行相等比较
                    let result = select_value.compare(BinaryOp::Eq, &case_value);
                    if let Value::Boolean(true) = result {
                        // 执行匹配的 Case body
                        for stmt in &case.body {
                            self.eval_stmt(stmt)?;
                        }
                        return Ok(Value::Empty);
                    }
                }
            }
        }

        // 如果没有匹配的 Case，执行 Else 块
        if let Some(else_stmts) = else_block {
            for stmt in else_stmts {
                self.eval_stmt(stmt)?;
            }
        }

        Ok(Value::Empty)
    }

    /// 注册 Sub
    fn eval_sub(
        &mut self,
        name: &str,
        params: &[crate::ast::Param],
        body: &[Stmt],
    ) -> Result<Value, RuntimeError> {
        self.context.functions.insert(
            name.to_lowercase(),
            Function {
                name: name.to_string(),
                params: params.iter().map(|p| p.name.clone()).collect(),
                body: body.to_vec(),
            },
        );
        Ok(Value::Empty)
    }

    /// 注册 Function
    fn eval_function(
        &mut self,
        name: &str,
        params: &[crate::ast::Param],
        body: &[Stmt],
    ) -> Result<Value, RuntimeError> {
        self.context.functions.insert(
            name.to_lowercase(),
            Function {
                name: name.to_string(),
                params: params.iter().map(|p| p.name.clone()).collect(),
                body: body.to_vec(),
            },
        );
        Ok(Value::Empty)
    }

    /// 执行 Call 语句
    fn eval_call(&mut self, name: &str, args: &[Expr]) -> Result<Value, RuntimeError> {
        let _args: Result<Vec<Value>, _> = args.iter().map(|e| self.eval_expr(e)).collect();
        // 简化：直接调用函数
        let name_lower = name.to_lowercase();
        if let Some(func) = self.context.functions.get(&name_lower).cloned() {
            self.context.push_scope();
            for stmt in &func.body {
                self.eval_stmt(stmt)?;
            }
            self.context.pop_scope();
        }
        Ok(Value::Empty)
    }

    /// Exit For
    fn eval_exit_for(&mut self) -> Result<Value, RuntimeError> {
        // TODO: 实现循环退出
        Ok(Value::Empty)
    }

    /// Exit Function
    fn eval_exit_function(&mut self) -> Result<Value, RuntimeError> {
        self.context.should_exit = true;
        Ok(Value::Empty)
    }

    /// Exit Sub
    fn eval_exit_sub(&mut self) -> Result<Value, RuntimeError> {
        self.context.should_exit = true;
        Ok(Value::Empty)
    }

    /// 求值表达式
    pub fn eval_expr(&mut self, expr: &Expr) -> Result<Value, RuntimeError> {
        match expr {
            Expr::Number(n) => Ok(Value::Number(*n)),
            Expr::String(s) => Ok(Value::String(s.clone())),
            Expr::Boolean(b) => Ok(Value::Boolean(*b)),
            Expr::Nothing => Ok(Value::Nothing),
            Expr::Empty => Ok(Value::Empty),
            Expr::Null => Ok(Value::Null),
            Expr::Variable(name) => self
                .context
                .get_var(name)
                .cloned()
                .ok_or_else(|| RuntimeError::UndefinedVariable(name.clone())),
            Expr::Binary { left, op, right } => {
                let left_val = self.eval_expr(left)?;
                let right_val = self.eval_expr(right)?;
                match op {
                    BinaryOp::Eq
                    | BinaryOp::Ne
                    | BinaryOp::Lt
                    | BinaryOp::Le
                    | BinaryOp::Gt
                    | BinaryOp::Ge
                    | BinaryOp::And
                    | BinaryOp::Or
                    | BinaryOp::Xor
                    | BinaryOp::Is => Ok(left_val.compare(*op, &right_val)),
                    _ => left_val.binary_op(*op, &right_val),
                }
            }
            Expr::Unary { op, operand } => {
                let val = self.eval_expr(operand)?;
                match op {
                    UnaryOp::Neg => Ok(Value::Number(-val.to_number())),
                    UnaryOp::Not => Ok(Value::Boolean(!val.is_truthy())),
                }
            }
            Expr::Call { name: _, args } => {
                let arg_values: Result<Vec<Value>, _> =
                    args.iter().map(|e| self.eval_expr(e)).collect();
                let _arg_values = arg_values?;
                // TODO: 实现函数调用
                Ok(Value::Empty)
            }
            Expr::Property { object, property } => {
                let _obj = self.eval_expr(object)?;
                // TODO: 实现属性访问
                Err(RuntimeError::PropertyNotFound(property.clone()))
            }
            Expr::Method { object, method, args } => {
                self.eval_method(object, method, args)
            }
            Expr::Array(elements) => {
                let values: Result<Vec<Value>, _> =
                    elements.iter().map(|e| self.eval_expr(e)).collect();
                Ok(Value::Array(values?))
            }
            Expr::Index { object, index } => {
                // 处理 Request("key") 语法
                let index_val = self.eval_expr(index)?;
                let index_key = match &index_val {
                    Value::String(s) => s.clone(),
                    _ => ValueConversion::to_string(&index_val),
                };

                match object.as_ref() {
                    // 特殊处理 Request 对象
                    Expr::Variable(name) if name.to_lowercase() == "request" => {
                        // 从 request_data 中获取值
                        match self.context.get_request_param(&index_key) {
                            Some(value) => Ok(Value::String(value.clone())),
                            None => Ok(Value::Empty),
                        }
                    }
                    // 处理数组或对象/字典访问，或内置函数调用
                    Expr::Variable(name) => {
                        let name_lower = name.to_lowercase();

                        // 检查是否是内置函数调用
                        if let Some(result) = self.call_builtin_function(&name_lower, &index_val) {
                            return result;
                        }

                        // 检查变量是否是数组或对象
                        if let Some(value) = self.context.get_var(name).cloned() {
                            match value {
                                Value::Array(arr) => {
                                    if let Value::Number(i) = index_val {
                                        let i = i as usize;
                                        if i < arr.len() {
                                            return Ok(arr[i].clone());
                                        }
                                    }
                                }
                                Value::Object(obj) => {
                                    if let Some(v) = obj.get(&index_key) {
                                        return Ok(v.clone());
                                    }
                                }
                                _ => {}
                            }
                        }
                        Err(RuntimeError::InvalidIndex)
                    }
                    _ => Err(RuntimeError::InvalidIndex),
                }
            }
            _ => Err(RuntimeError::Generic(format!(
                "Unimplemented expr: {:?}",
                expr
            ))),
        }
    }

    /// 调用内置函数（用于处理类似 CInt(x) 的调用）
    fn call_builtin_function(&mut self, name: &str, arg: &Value) -> Option<Result<Value, RuntimeError>> {
        match name {
            // 类型转换函数
            "cint" | "cbyte" | "cbool" => {
                use crate::runtime::ValueConversion;
                Some(Ok(Value::Number(ValueConversion::to_number(arg) as i32 as f64)))
            }
            "clng" | "csng" => {
                use crate::runtime::ValueConversion;
                Some(Ok(Value::Number(ValueConversion::to_number(arg))))
            }
            "cdbl" => {
                use crate::runtime::ValueConversion;
                Some(Ok(Value::Number(ValueConversion::to_number(arg))))
            }
            "cstr" => {
                use crate::runtime::ValueConversion;
                Some(Ok(Value::String(ValueConversion::to_string(arg))))
            }
            "cdate" => {
                // TODO: 实现日期转换
                Some(Ok(arg.clone()))
            }
            "int" | "fix" => {
                use crate::runtime::ValueConversion;
                let n = ValueConversion::to_number(arg);
                Some(Ok(Value::Number(n.trunc())))
            }
            "abs" => {
                use crate::runtime::ValueConversion;
                let n = ValueConversion::to_number(arg);
                Some(Ok(Value::Number(n.abs())))
            }
            "sgn" => {
                use crate::runtime::ValueConversion;
                let n = ValueConversion::to_number(arg);
                Some(Ok(Value::Number(if n > 0.0 { 1.0 } else if n < 0.0 { -1.0 } else { 0.0 })))
            }
            "sqr" => {
                use crate::runtime::ValueConversion;
                let n = ValueConversion::to_number(arg);
                Some(Ok(Value::Number(n.sqrt())))
            }
            "len" => {
                use crate::runtime::ValueConversion;
                let s = ValueConversion::to_string(arg);
                Some(Ok(Value::Number(s.len() as f64)))
            }
            "trim" | "ltrim" | "rtrim" => {
                use crate::runtime::ValueConversion;
                let s = ValueConversion::to_string(arg);
                let result = match name {
                    "trim" => s.trim().to_string(),
                    "ltrim" => s.trim_start().to_string(),
                    "rtrim" => s.trim_end().to_string(),
                    _ => s,
                };
                Some(Ok(Value::String(result)))
            }
            "ucase" | "lcase" => {
                use crate::runtime::ValueConversion;
                let s = ValueConversion::to_string(arg);
                let result = match name {
                    "ucase" => s.to_uppercase(),
                    "lcase" => s.to_lowercase(),
                    _ => s,
                };
                Some(Ok(Value::String(result)))
            }
            "chr" => {
                use crate::runtime::ValueConversion;
                let n = ValueConversion::to_number(arg) as u32;
                Some(Ok(Value::String(char::from_u32(n).unwrap_or('\0').to_string())))
            }
            "asc" => {
                use crate::runtime::ValueConversion;
                let s = ValueConversion::to_string(arg);
                let code = s.chars().next().map(|c| c as u8 as f64).unwrap_or(0.0);
                Some(Ok(Value::Number(code)))
            }
            "isnumeric" => {
                // 检查值是否可以转换为数字
                let is_num = match arg {
                    Value::Number(_) => true,
                    Value::Boolean(_) => true,
                    Value::String(s) => s.parse::<f64>().is_ok(),
                    Value::Empty => true,
                    Value::Null => false,
                    Value::Nothing => false,
                    Value::Array(_) => false,
                    Value::Object(_) => false,
                };
                Some(Ok(Value::Boolean(is_num)))
            }
            "isempty" => {
                Some(Ok(Value::Boolean(matches!(arg, Value::Empty))))
            }
            "isnull" => {
                Some(Ok(Value::Boolean(matches!(arg, Value::Null))))
            }
            "isarray" => {
                Some(Ok(Value::Boolean(matches!(arg, Value::Array(_)))))
            }
            "isobject" => {
                Some(Ok(Value::Boolean(matches!(arg, Value::Object(_)))))
            }
            "isdate" => {
                // TODO: 实现日期检测
                Some(Ok(Value::Boolean(false)))
            }
            _ => None,
        }
    }

    /// 执行方法调用
    fn eval_method(
        &mut self,
        object: &Expr,
        method: &str,
        args: &[Expr],
    ) -> Result<Value, RuntimeError> {
        // 获取对象名称（如果是变量）
        let object_name = match object {
            Expr::Variable(name) => Some(name.to_lowercase()),
            _ => None,
        };

        let method_lower = method.to_lowercase();

        // 处理内建对象的方法
        match (object_name.as_deref(), method_lower.as_str()) {
            // Response.Write
            (Some("response"), "write") => {
                if !args.is_empty() {
                    let value = self.eval_expr(&args[0])?;
                    use crate::runtime::ValueConversion;
                    self.context.write(&ValueConversion::to_string(&value));
                }
                Ok(Value::Empty)
            }
            // Response.End
            (Some("response"), "end") => {
                // TODO: 实现响应结束
                Ok(Value::Empty)
            }
            // Response.Redirect
            (Some("response"), "redirect") => {
                // TODO: 实现重定向
                Ok(Value::Empty)
            }
            // Request.QueryString / Request.Form
            (Some("request"), "querystring") => {
                if !args.is_empty() {
                    let key = self.eval_expr(&args[0])?;
                    use crate::runtime::ValueConversion;
                    let key_str = ValueConversion::to_string(&key);
                    // 从上下文获取 QueryString
                    Ok(self.context.get_var(&key_str).cloned().unwrap_or(Value::Empty))
                } else {
                    Ok(Value::Empty)
                }
            }
            (Some("request"), "form") => {
                if !args.is_empty() {
                    let key = self.eval_expr(&args[0])?;
                    use crate::runtime::ValueConversion;
                    let key_str = ValueConversion::to_string(&key);
                    // 从上下文获取 Form 数据
                    Ok(self.context.get_var(&key_str).cloned().unwrap_or(Value::Empty))
                } else {
                    Ok(Value::Empty)
                }
            }
            // Server.CreateObject
            (Some("server"), "createobject") => {
                // 不支持 COM 对象创建
                Err(RuntimeError::Generic(
                    "COM object creation is not supported".to_string(),
                ))
            }
            // 其他方法调用
            _ => {
                // 尝试调用用户定义的方法
                let arg_values: Result<Vec<Value>, _> =
                    args.iter().map(|e| self.eval_expr(e)).collect();
                let _arg_values = arg_values?;
                // TODO: 实现用户定义方法调用
                Ok(Value::Empty)
            }
        }
    }
}

impl Default for Interpreter {
    fn default() -> Self {
        Self::new()
    }
}
