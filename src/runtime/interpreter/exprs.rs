//! 表达式求值模块
//!
//! 处理各种 VBScript 表达式的求值逻辑

use crate::ast::{BinaryOp, Expr, UnaryOp};
use crate::runtime::{BuiltinObject, RuntimeError, Value, ValueCompare, ValueConversion, ValueOps, VbsClass};

use super::Interpreter;

/// 表达式求值器
impl Interpreter {
    /// 求值表达式
    pub fn eval_expr(&mut self, expr: &Expr) -> Result<Value, RuntimeError> {
        match expr {
            Expr::Number(n) => Ok(Value::Number(*n)),
            Expr::String(s) => Ok(Value::String(s.clone())),
            Expr::Boolean(b) => Ok(Value::Boolean(*b)),
            Expr::Date(date_str) => {
                // 日期字面量暂时作为字符串处理
                // 后续可以在需要时转换为实际的日期值
                Ok(Value::String(date_str.clone()))
            }
            Expr::Nothing => Ok(Value::Nothing),
            Expr::Empty => Ok(Value::Empty),
            Expr::Null => Ok(Value::Null),
            Expr::Variable(name) => {
                // 首先尝试从上下文中获取变量
                if let Some(value) = self.context.get_var(name).cloned() {
                    Ok(value)
                } else {
                    // 如果没有找到变量，检查是否是内置函数（无参数调用）
                    // 支持：Now、Date、Time 等无参数内置函数
                    if let Some(result) = self.try_call_builtin(name, &[]) {
                        return result;
                    }
                    // 检查是否是用户定义的函数（无参数调用）
                    if self.context.get_function(name).is_some() {
                        self.eval_user_function_call(name, &[])
                    } else {
                        Ok(Value::Empty)
                    }
                }
            }
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
                self.eval_call_expr(name, args)
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
            Expr::New(class_name) => {
                self.eval_new(class_name)
            }
            Expr::Index { object, index } => self.eval_index(object, index),
            // 这个分支理论上不可达，因为所有 Expr 变体都已处理
            // 保留它以防止将来添加新变体时忘记处理
            _ => Ok(Value::Empty),
        }
    }

    /// 尝试调用内置函数
    fn try_call_builtin(&mut self, name: &str, args: &[Value]) -> Option<Result<Value, RuntimeError>> {
        use crate::runtime::builtins::{TokenRegistry, BuiltinExecutor};

        let registry = TokenRegistry::new();
        registry.lookup(name).map(|token| BuiltinExecutor::execute(token, args))
    }

    fn eval_call_expr(&mut self, name: &str, args: &[Expr]) -> Result<Value, RuntimeError> {
        let arg_values: Result<Vec<Value>, _> = args.iter().map(|e| self.eval_expr(e)).collect();
        let arg_values = arg_values?;

        eprintln!("DEBUG: Call {} with {:?} args", name, arg_values.len());

        // 尝试调用内置函数
        if let Some(result) = self.try_call_builtin(name, &arg_values) {
            return result;
        }

        self.eval_user_function_call(name, &arg_values)
    }

    fn eval_user_function_call(&mut self, name: &str, arg_values: &[Value]) -> Result<Value, RuntimeError> {
        if let Some(func) = self.context.get_function(name).cloned() {
            eprintln!("DEBUG: Found user function {}", name);
            self.context.push_scope();

            for (i, param_name) in func.params.iter().enumerate() {
                let value = if i < arg_values.len() {
                    arg_values[i].clone()
                } else {
                    Value::Empty
                };
                self.context.define_var(param_name.clone(), value);
            }

            for stmt in &func.body {
                match self.eval_stmt(stmt) {
                    Ok(_) => {}
                    Err(e) => {
                        self.context.pop_scope();
                        return Err(e);
                    }
                }
            }

            let func_name_lower = crate::utils::normalize_identifier(&func.name);
            let result = self.context.get_var(&func.name)
                .or_else(|| self.context.get_var(&func_name_lower))
                .cloned()
                .unwrap_or(Value::Empty);

            self.context.pop_scope();
            return Ok(result);
        }

        if arg_values.len() == 1 {
            // 处理数组索引访问：arr(0)
            if let Some(Value::Array(arr)) = self.context.get_var(name) {
                if let Value::Number(i) = &arg_values[0] {
                    let i = *i as usize;
                    if i < arr.len() {
                        return Ok(arr[i].clone());
                    }
                }
            }
        }

        Err(RuntimeError::UndefinedVariable(format!("Function '{}' or array index", name)))
    }

    /// 处理索引表达式
    fn eval_index(&mut self, object: &Expr, index: &Expr) -> Result<Value, RuntimeError> {
        let index_val = self.eval_expr(index)?;

        match object {
            // 处理方法调用后的索引访问，如 Request.Form("name")(1)
            Expr::Method { .. } | Expr::Property { .. } => {
                let obj_val = self.eval_expr(object)?;
                self.eval_index_on_value(&obj_val, &index_val)
            }
            // 处理变量索引访问
            Expr::Variable(name) => {
                // 检查是否是内置函数调用
                if let Some(result) = self.try_call_builtin(name, &[index_val.clone()]) {
                    return result;
                }

                // 检查变量是否是数组或对象
                if let Some(value) = self.context.get_var(name).cloned() {
                    self.eval_index_on_value(&value, &index_val)
                } else {
                    Err(RuntimeError::InvalidIndex)
                }
            }
            _ => {
                // 先求值 object，再进行索引访问
                let obj_val = self.eval_expr(object)?;
                self.eval_index_on_value(&obj_val, &index_val)
            }
        }
    }

    /// 对 Value 进行索引访问
    fn eval_index_on_value(&self, value: &Value, index: &Value) -> Result<Value, RuntimeError> {
        match value {
            Value::Array(arr) => {
                if let Value::Number(i) = index {
                    let i = *i as usize;
                    // ASP 中索引从 1 开始
                    if i >= 1 && i <= arr.len() {
                        return Ok(arr[i - 1].clone());
                    }
                }
                Ok(Value::Empty)
            }
            Value::String(s) => {
                // ASP 中字符串的索引访问：对于单值，(1) 返回字符串本身
                if let Value::Number(i) = index {
                    if *i == 1.0 {
                        return Ok(Value::String(s.clone()));
                    }
                }
                Ok(Value::Empty)
            }
            Value::Object(obj) => {
                // 使用 BuiltinObject trait 的 index 方法
                match obj.index(index) {
                    Ok(value) => Ok(value),
                    Err(_) => Ok(Value::Empty),
                }
            }
            _ => Ok(Value::Empty),
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
            // Response 方法 - 统一通过 BuiltinObject trait 调用
            (Some("response"), method_name) => {
                // 计算参数
                let arg_values: Result<Vec<Value>, _> =
                    args.iter().map(|e| self.eval_expr(e)).collect();
                let arg_values = arg_values?;

                // 特殊处理 Response.End - 设置退出标志
                if method_name.to_lowercase() == "end" {
                    self.context.response_mut().end();
                    self.context.should_exit = true;
                    return Ok(Value::Empty);
                }

                // 调用 Response 对象的方法
                let response = self.context.response_mut();
                response.call_method(method_name, arg_values)
            }
            // Request.QueryString / Request.Form
            (Some("request"), "querystring" | "form") => {
                self.eval_request_method(args)
            }
            // Server.CreateObject
            (Some("server"), "createobject") => {
                // 不支持 COM 对象创建，返回 Empty 以避免页面中断
                Ok(Value::Empty)
            }
            // 其他方法调用
            _ => {
                // 尝试调用用户定义的方法
                let arg_values: Result<Vec<Value>, _> =
                    args.iter().map(|e| self.eval_expr(e)).collect();
                let arg_values = arg_values?;

                // 特殊处理 Session.Contents 的方法调用
                // 当调用 Session.Contents.Remove(key) 时
                if let Expr::Property { object, property: contents_prop } = object {
                    if contents_prop.to_lowercase() == "contents" {
                        if let Expr::Variable(session_name) = object.as_ref() {
                            if session_name.to_lowercase() == "session" {
                                // 这是 Session.Contents 的方法调用
                                // 通过 Session 对象的 call_method 来处理
                                if let Some(Value::Object(mut session_obj)) = self.context.get_var("Session").cloned() {
                                    let result = session_obj.call_method(method, arg_values);
                                    // 将修改后的 session 对象写回
                                    self.context.set_var("Session".to_string(), Value::Object(session_obj));
                                    return result;
                                }
                            }
                        }
                    }
                }

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
            Some("request") => self.eval_request_property(property_lower.as_str()),
            Some("response") => {
                // 处理无参数方法调用：Response.Clear 和 Response.End
                match property_lower.as_str() {
                    "clear" => {
                        self.context.response_mut().clear();
                        return Ok(Value::Empty);
                    }
                    "end" => {
                        self.context.response_mut().end();
                        // 设置退出标志，停止后续执行
                        self.context.should_exit = true;
                        return Ok(Value::Empty);
                    }
                    _ => {}
                }
                // 从 response 对象获取属性
                let response = self.context.response();
                match response.get_property(property) {
                    Ok(val) => return Ok(val),
                    Err(_) => {
                        // 如果 BuiltinObject 返回错误，提供默认空值
                        match property_lower.as_str() {
                            "status" | "contenttype" | "buffer" | "charset" | "cachecontrol"
                            | "expires" | "expiresabsolute" | "pics" | "isclientconnected" | "cookies" => Ok(Value::Empty),
                            _ => Err(RuntimeError::PropertyNotFound(property.to_string())),
                        }
                    }
                }
            }
            Some("server") => {
                match property_lower.as_str() {
                    "scripttimeout" => Ok(Value::Number(90.0)),
                    _ => Err(RuntimeError::PropertyNotFound(property.to_string())),
                }
            }
            Some("session") => {
                // 处理 Session 对象的属性访问
                match property_lower.as_str() {
                    "contents" => {
                        // 返回特殊的 SessionContents 标记对象
                        let mut contents_obj = std::collections::HashMap::new();
                        contents_obj.insert("__session_contents__".to_string(), Value::Boolean(true));
                        contents_obj.insert("count".to_string(), Value::Number(-1.0));
                        return Ok(Value::from_hashmap(contents_obj));
                    }
                    _ => {
                        // 处理其他 Session 属性
                        if let Some(Value::Object(session_obj)) = self.context.get_var("Session").cloned() {
                            // 使用 BuiltinObject trait 的 get_property 方法
                            match session_obj.get_property(property) {
                                Ok(value) => return Ok(value),
                                Err(_) => {}
                            }
                        }
                        Err(RuntimeError::PropertyNotFound(property.to_string()))
                    }
                }
            }
            _ => {
                // 处理通用属性访问（从变量或表达式中获取对象）
                let obj_value = self.eval_expr(object)?;

                // 处理 .Count 属性（适用于各种类型）
                if property_lower == "count" {
                    return Ok(match &obj_value {
                        Value::Array(arr) => Value::Number(arr.len() as f64),
                        Value::Object(obj) => {
                            // 使用 BuiltinObject 的 get_property 方法
                            obj.get_property("count").unwrap_or(Value::Number(0.0))
                        }
                        Value::String(_) | Value::Number(_) | Value::Boolean(_)
                        | Value::Empty | Value::Null | Value::Nothing => Value::Number(1.0),
                    });
                }

                // 处理对象的属性访问
                match obj_value {
                    Value::Object(obj) => {
                        // 使用 BuiltinObject trait 的 get_property 方法
                        match obj.get_property(property) {
                            Ok(value) => return Ok(value),
                            Err(_) => {}
                        }
                    }
                    _ => {}
                }

                Err(RuntimeError::PropertyNotFound(property.to_string()))
            }
        }
    }

    fn eval_request_method(&mut self, args: &[Expr]) -> Result<Value, RuntimeError> {
        if args.is_empty() {
            return Ok(Value::Empty);
        }
        let key = self.eval_expr(&args[0])?;
        let key_str = ValueConversion::to_string(&key);
        Ok(self.context.get_request_param_all(&key_str)
            .map(|values| {
                if values.len() == 1 {
                    Value::String(values[0].clone())
                } else {
                    Value::Array(values.iter().map(|s| Value::String(s.clone())).collect())
                }
            })
            .unwrap_or(Value::Empty))
    }

    fn eval_request_property(&self, property: &str) -> Result<Value, RuntimeError> {
        match property {
            "form" | "querystring" => {
                let mut data = std::collections::HashMap::new();
                for (key, values) in self.context.request_data.iter() {
                    let value_str = if values.len() == 1 {
                        values[0].clone()
                    } else {
                        values.join(", ")
                    };
                    data.insert(key.clone(), Value::String(value_str));
                }
                Ok(Value::from_hashmap(data))
            }
            "cookies" | "servervariables" => {
                Ok(Value::new_dictionary())
            }
            _ => Err(RuntimeError::PropertyNotFound(property.to_string())),
        }
    }

    /// 执行 New 表达式 - 创建类实例
    fn eval_new(&mut self, class_name: &str) -> Result<Value, RuntimeError> {
        let normalized_name = crate::utils::normalize_identifier(class_name);

        // 从上下文中查找类定义
        if let Some(class_def) = self.context.classes.get(&normalized_name) {
            // 创建 VbsClass
            let vbs_class = VbsClass::from_ast(class_def.name.clone(), class_def.members.clone());

            // 创建实例
            let instance = vbs_class.new_instance();

            // 返回实例的 Value 表示
            Ok(instance.to_value())
        } else {
            Err(RuntimeError::Generic(format!(
                "Class '{}' not found",
                class_name
            )))
        }
    }

    #[allow(dead_code)]
    fn is_session_user_key(k: &&str) -> bool {
        !k.starts_with("__") && *k != "sessionid" && *k != "timeout"
    }
}
