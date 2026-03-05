//! 表达式求值模块
//!
//! 处理各种 VBScript 表达式的求值逻辑

use crate::ast::{BinaryOp, Expr, UnaryOp};
use crate::runtime::{RuntimeError, Value, ValueCompare, ValueConversion, ValueOps};
use crate::utils::identifier_matches;

use super::Interpreter;

/// 表达式求值器
impl Interpreter {
    /// 求值表达式
    pub fn eval_expr(&mut self, expr: &Expr) -> Result<Value, RuntimeError> {
        match expr {
            Expr::Number(n) => Ok(Value::Number(*n)),
            Expr::String(s) => Ok(Value::String(s.clone())),
            Expr::Boolean(b) => Ok(Value::Boolean(*b)),
            Expr::Nothing => Ok(Value::Nothing),
            Expr::Empty => Ok(Value::Empty),
            Expr::Null => Ok(Value::Null),
            Expr::Variable(name) => Ok(self
                .context
                .get_var(name)
                .cloned()
                .unwrap_or(Value::Empty)),
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
            Expr::Call { name, args } => {
                // 计算参数值
                let arg_values: Result<Vec<Value>, _> =
                    args.iter().map(|e| self.eval_expr(e)).collect();
                let arg_values = arg_values?;

                eprintln!("DEBUG: Call {} with {:?} args", name, arg_values.len());

                // 首先尝试作为内置函数调用
                if let Some(result) = super::builtins::call_builtin_function_multi(name, &arg_values) {
                    return result;
                }

                // 单参数内置函数回退
                if arg_values.len() == 1 {
                    if let Some(result) =
                        super::builtins::call_builtin_function(&name.to_lowercase(), &arg_values[0])
                    {
                        return result;
                    }
                }

                // 尝试作为用户定义函数调用
                if let Some(func) = self.context.get_function(name).cloned() {
                    eprintln!("DEBUG: Found user function {}", name);

                    // 创建新的作用域
                    self.context.push_scope();

                    // 绑定参数到函数作用域
                    for (i, param_name) in func.params.iter().enumerate() {
                        let value = if i < arg_values.len() {
                            arg_values[i].clone()
                        } else {
                            Value::Empty
                        };
                        self.context.define_var(param_name.clone(), value);
                    }

                    // 执行函数体
                    for stmt in &func.body {
                        match self.eval_stmt(stmt) {
                            Ok(_) => {}
                            Err(e) => {
                                self.context.pop_scope();
                                return Err(e);
                            }
                        }
                    }

                    // 获取函数返回值（在 VBScript 中，函数名作为返回值变量）
                    let func_name_lower = crate::utils::normalize_identifier(&func.name);
                    let result = self.context.get_var(&func.name)
                        .or_else(|| self.context.get_var(&func_name_lower))
                        .cloned()
                        .unwrap_or(Value::Empty);

                    // 弹出作用域
                    self.context.pop_scope();

                    return Ok(result);
                }

                // 未找到函数，尝试作为数组索引访问
                // 支持 fruits(i) 这样的数组访问语法
                if arg_values.len() == 1 {
                    eprintln!("DEBUG: Checking if {} is an array variable", name);
                    if let Some(var_val) = self.context.get_var(name) {
                        eprintln!("DEBUG: Found variable {} with value {:?}", name, var_val);
                    }
                    if let Some(Value::Array(arr)) = self.context.get_var(name) {
                        eprintln!("DEBUG: Found array {} with length {}", name, arr.len());
                        let idx = &arg_values[0];
                        if let Value::Number(i) = idx {
                            let i = *i as usize;
                            if i < arr.len() {
                                return Ok(arr[i].clone());
                            }
                        }
                    }
                }

                // 未找到函数或索引
                Err(RuntimeError::UndefinedVariable(format!("Function '{}' or array index", name)))
            }
            Expr::Property { object, property } => {
                self.eval_property(object, property)
            }
            Expr::Method { object, method, args } => {
                self.eval_method(object, method, args)
            }
            Expr::Array(elements) => {
                let values: Result<Vec<Value>, _> =
                    elements.iter().map(|e| self.eval_expr(e)).collect();
                Ok(Value::Array(values?))
            }
            Expr::Index { object, index } => self.eval_index(object, index),
            _ => Err(RuntimeError::Generic(format!(
                "Unimplemented expr: {:?}",
                expr
            ))),
        }
    }

    /// 处理索引表达式
    fn eval_index(&mut self, object: &Expr, index: &Expr) -> Result<Value, RuntimeError> {
        let index_val = self.eval_expr(index)?;
        let index_key = match &index_val {
            Value::String(s) => s.clone(),
            _ => ValueConversion::to_string(&index_val),
        };

        match object {
            // 特殊处理 Request 对象
            Expr::Variable(name) if identifier_matches(name, "request") => {
                // 从 request_data 中获取值
                match self.context.get_request_param(&index_key) {
                    Some(value) => Ok(Value::String(value.clone())),
                    None => Ok(Value::Empty),
                }
            }
            // 处理数组或对象/字典访问，或内置函数调用
            Expr::Variable(name) => {
                // 检查是否是内置函数调用
                if let Some(result) =
                    super::builtins::call_builtin_function(&name.to_lowercase(), &index_val)
                {
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
            // 处理方法调用后的索引访问，如 Request.Form("name")(1)
            Expr::Method { .. } | Expr::Property { .. } => {
                // 先求值 object
                let obj_val = self.eval_expr(object)?;

                // 根据索引值的类型进行访问
                match obj_val {
                    Value::Array(arr) => {
                        if let Value::Number(i) = index_val {
                            let i = i as usize;
                            // ASP 中索引从 1 开始
                            if i >= 1 && i <= arr.len() {
                                return Ok(arr[i - 1].clone());
                            }
                        }
                        Ok(Value::Empty)
                    }
                    Value::String(s) => {
                        // ASP 中字符串的索引访问：对于单值，(1) 返回字符串本身
                        if let Value::Number(i) = index_val {
                            if i == 1.0 {
                                return Ok(Value::String(s));
                            }
                        }
                        Ok(Value::Empty)
                    }
                    Value::Object(obj) => {
                        if let Some(v) = obj.get(&index_key) {
                            return Ok(v.clone());
                        }
                        Ok(Value::Empty)
                    }
                    _ => Ok(Value::Empty),
                }
            }
            _ => Err(RuntimeError::InvalidIndex),
        }
    }

    /// 执行方法调用
    pub fn eval_method(
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
                    let text = ValueConversion::to_string(&value);
                    // 同时写入 Context 的 output 和 Response 对象的 buffer
                    self.context.write(&text);
                    self.context.response_mut().write(&text);
                }
                Ok(Value::Empty)
            }
            // Response.End
            (Some("response"), "end") => {
                self.context.response_mut().end();
                self.context.set_should_exit();
                Ok(Value::Empty)
            }
            // Response.Redirect
            (Some("response"), "redirect") => {
                if !args.is_empty() {
                    let url = self.eval_expr(&args[0])?;
                    let url_str = ValueConversion::to_string(&url);
                    self.context.response_mut().redirect(&url_str);
                    self.context.set_should_exit();
                }
                Ok(Value::Empty)
            }
            // Response.Clear
            (Some("response"), "clear") => {
                self.context.response_mut().clear();
                self.context.clear_output();
                Ok(Value::Empty)
            }
            // Response.AddHeader
            (Some("response"), "addheader") => {
                if args.len() >= 2 {
                    let name = self.eval_expr(&args[0])?;
                    let value = self.eval_expr(&args[1])?;
                    let name_str = ValueConversion::to_string(&name);
                    let value_str = ValueConversion::to_string(&value);
                    self.context.response_mut().add_header(&name_str, &value_str);
                }
                Ok(Value::Empty)
            }
            // Request.QueryString / Request.Form
            (Some("request"), "querystring") => {
                if !args.is_empty() {
                    let key = self.eval_expr(&args[0])?;
                    let key_str = ValueConversion::to_string(&key);
                    // 从 request_data 获取 QueryString（智能处理多值）
                    Ok(self.context.get_request_param_all(&key_str)
                        .map(|values| {
                            if values.len() == 1 {
                                // 单值：返回字符串
                                Value::String(values[0].clone())
                            } else {
                                // 多值：返回数组
                                Value::Array(values.iter().map(|s| Value::String(s.clone())).collect())
                            }
                        })
                        .unwrap_or(Value::Empty))
                } else {
                    Ok(Value::Empty)
                }
            }
            (Some("request"), "form") => {
                if !args.is_empty() {
                    let key = self.eval_expr(&args[0])?;
                    let key_str = ValueConversion::to_string(&key);
                    // 从 request_data 获取 Form 数据（智能处理多值）
                    Ok(self.context.get_request_param_all(&key_str)
                        .map(|values| {
                            if values.len() == 1 {
                                // 单值：返回字符串
                                Value::String(values[0].clone())
                            } else {
                                // 多值：返回数组
                                Value::Array(values.iter().map(|s| Value::String(s.clone())).collect())
                            }
                        })
                        .unwrap_or(Value::Empty))
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

    /// 处理属性访问表达式
    fn eval_property(&mut self, object: &Expr, property: &str) -> Result<Value, RuntimeError> {
        // 获取对象名称（如果是变量）
        let object_name = match object {
            Expr::Variable(name) => Some(name.to_lowercase()),
            _ => None,
        };

        let property_lower = property.to_lowercase();

        // 处理内建对象的属性访问
        match object_name.as_deref() {
            Some("request") => {
                match property_lower.as_str() {
                    "form" => {
                        // 返回表单数据集合（多值转为逗号分隔字符串）
                        let mut form_data = std::collections::HashMap::new();
                        for (key, values) in self.context.request_data.iter() {
                            let value_str = if values.len() == 1 {
                                values[0].clone()
                            } else {
                                values.join(", ")
                            };
                            form_data.insert(key.clone(), Value::String(value_str));
                        }
                        Ok(Value::Object(form_data))
                    }
                    "querystring" => {
                        // 返回查询字符串集合（多值转为逗号分隔字符串）
                        let mut query_data = std::collections::HashMap::new();
                        for (key, values) in self.context.request_data.iter() {
                            let value_str = if values.len() == 1 {
                                values[0].clone()
                            } else {
                                values.join(", ")
                            };
                            query_data.insert(key.clone(), Value::String(value_str));
                        }
                        Ok(Value::Object(query_data))
                    }
                    "cookies" | "servervariables" => {
                        // 返回空对象（暂不支持）
                        Ok(Value::Object(std::collections::HashMap::new()))
                    }
                    _ => Err(RuntimeError::PropertyNotFound(property.to_string())),
                }
            }
            Some("response") => {
                match property_lower.as_str() {
                    "status" | "contenttype" => Ok(Value::Empty),
                    _ => Err(RuntimeError::PropertyNotFound(property.to_string())),
                }
            }
            Some("server") => {
                match property_lower.as_str() {
                    "scripttimeout" => Ok(Value::Number(90.0)),
                    _ => Err(RuntimeError::PropertyNotFound(property.to_string())),
                }
            }
            _ => {
                // 处理通用属性访问（从变量或表达式中获取对象）
                let obj_value = self.eval_expr(object)?;

                // 处理 .Count 属性（适用于各种类型）
                if property_lower == "count" {
                    return Ok(match &obj_value {
                        Value::Array(arr) => Value::Number(arr.len() as f64),
                        Value::Object(obj) => Value::Number(obj.len() as f64),
                        Value::String(_) | Value::Number(_) | Value::Boolean(_)
                        | Value::Empty | Value::Null | Value::Nothing => Value::Number(1.0),
                    });
                }

                // 处理对象的属性访问
                match obj_value {
                    Value::Object(obj) => {
                        if let Some(v) = obj.get(&property_lower) {
                            return Ok(v.clone());
                        }
                    }
                    _ => {}
                }

                Err(RuntimeError::PropertyNotFound(property.to_string()))
            }
        }
    }
}
