//! 变量作用域

use super::Value;
use std::collections::HashMap;

/// 作用域
#[derive(Debug, Clone)]
pub struct Scope {
    /// 变量
    pub variables: HashMap<String, Value>,
    /// 父作用域
    pub parent: Option<Box<Scope>>,
}

impl Scope {
    /// 创建新作用域
    pub fn new() -> Self {
        Scope {
            variables: HashMap::new(),
            parent: None,
        }
    }

    /// 创建带父作用域的子作用域
    pub fn with_parent(parent: Scope) -> Self {
        Scope {
            variables: HashMap::new(),
            parent: Some(Box::new(parent)),
        }
    }

    /// 获取变量
    pub fn get(&self, name: &str) -> Option<&Value> {
        let name_lower = name.to_lowercase();
        self.variables
            .get(&name_lower)
            .or_else(|| self.parent.as_ref().and_then(|p| p.get(&name_lower)))
    }

    /// 设置变量
    pub fn set(&mut self, name: String, value: Value) {
        let name_lower = name.to_lowercase();
        self.variables.insert(name_lower, value);
    }

    /// 检查变量是否存在
    pub fn contains(&self, name: &str) -> bool {
        let name_lower = name.to_lowercase();
        self.variables.contains_key(&name_lower)
            || self
                .parent
                .as_ref()
                .map_or(false, |p| p.contains(&name_lower))
    }

    /// 在当前作用域定义变量
    pub fn define(&mut self, name: String, value: Value) {
        let name_lower = name.to_lowercase();
        self.variables.insert(name_lower, value);
    }
}

impl Default for Scope {
    fn default() -> Self {
        Self::new()
    }
}
