//! VBScript 运行时解释器

use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use thiserror::Error;

/// VBScript 值类型
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Value {
    Empty,
    Null,
    Nothing,
    Boolean(bool),
    Number(f64),
    String(String),
    Array(Vec<Value>),
    Object(HashMap<String, Value>),
}

impl Default for Value {
    fn default() -> Self {
        Value::Empty
    }
}

impl std::fmt::Display for Value {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Value::Empty => write!(f, ""),
            Value::Null => write!(f, "Null"),
            Value::Nothing => write!(f, "Nothing"),
            Value::Boolean(b) => write!(f, "{}", b),
            Value::Number(n) => write!(f, "{}", n),
            Value::String(s) => write!(f, "{}", s),
            Value::Array(arr) => {
                let items: Vec<String> = arr.iter().map(|v| v.to_string()).collect();
                write!(f, "[{}]", items.join(", "))
            }
            Value::Object(obj) => {
                let items: Vec<String> = obj.iter().map(|(k, v)| format!("{}: {}", k, v)).collect();
                write!(f, "{{{}}}", items.join(", "))
            }
        }
    }
}

/// 运行时错误
#[derive(Debug, Error)]
pub enum RuntimeError {
    #[error("Undefined variable: {0}")]
    UndefinedVariable(String),

    #[error("Type mismatch: {0}")]
    TypeMismatch(String),

    #[error("Division by zero")]
    DivisionByZero,

    #[error("Index out of bounds: {0}")]
    IndexOutOfBounds(usize),

    #[error("Object required")]
    ObjectRequired,

    #[error("Method not found: {0}")]
    MethodNotFound(String),

    #[error("Property not found: {0}")]
    PropertyNotFound(String),

    #[error("Argument count mismatch")]
    ArgumentCountMismatch,

    #[error("Runtime error: {0}")]
    Generic(String),
}

/// 变量作用域
#[derive(Debug, Clone)]
pub struct Scope {
    pub variables: HashMap<String, Value>,
    pub parent: Option<Box<Scope>>,
}

impl Scope {
    pub fn new() -> Self {
        Scope {
            variables: HashMap::new(),
            parent: None,
        }
    }

    pub fn with_parent(parent: Scope) -> Self {
        Scope {
            variables: HashMap::new(),
            parent: Some(Box::new(parent)),
        }
    }

    pub fn get(&self, name: &str) -> Option<&Value> {
        self.variables.get(name).or_else(|| {
            self.parent.as_ref().and_then(|p| p.get(name))
        })
    }

    pub fn set(&mut self, name: String, value: Value) {
        self.variables.insert(name, value);
    }
}

/// VBScript 解释器
pub struct Interpreter {
    global_scope: Scope,
    /// 内置对象（如 Response, Request 等）
    builtins: HashMap<String, Box<dyn BuiltinObject>>,
}

/// 内置对象 trait
pub trait BuiltinObject: Send + Sync {
    fn get_property(&self, name: &str) -> Result<Value, RuntimeError>;
    fn set_property(&mut self, name: &str, value: Value) -> Result<(), RuntimeError>;
    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError>;
}

impl Interpreter {
    pub fn new() -> Self {
        Interpreter {
            global_scope: Scope::new(),
            builtins: HashMap::new(),
        }
    }

    /// 注册内置对象
    pub fn register_builtin(&mut self, name: &str, obj: Box<dyn BuiltinObject>) {
        self.builtins.insert(name.to_lowercase(), obj);
    }

    /// 执行 AST
    pub fn execute(&mut self, ast: &[super::parser::Ast]) -> Result<Value, RuntimeError> {
        let mut result = Value::Empty;
        for stmt in ast {
            result = self.execute_stmt(stmt)?;
        }
        Ok(result)
    }

    fn execute_stmt(&mut self, stmt: &super::parser::Ast) -> Result<Value, RuntimeError> {
        match stmt {
            super::parser::Ast::String(s) => Ok(Value::String(s.clone())),
            super::parser::Ast::Number(n) => Ok(Value::Number(*n)),
            super::parser::Ast::Boolean(b) => Ok(Value::Boolean(*b)),
            super::parser::Ast::Ident(name) => {
                self.global_scope.get(name)
                    .cloned()
                    .ok_or_else(|| RuntimeError::UndefinedVariable(name.clone()))
            }
            _ => Err(RuntimeError::Generic(format!("Unimplemented: {:?}", stmt))),
        }
    }
}

impl Default for Interpreter {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_value_display() {
        assert_eq!(Value::String("hello".to_string()).to_string(), "hello");
        assert_eq!(Value::Number(42.0).to_string(), "42");
        assert_eq!(Value::Boolean(true).to_string(), "true");
    }
}
