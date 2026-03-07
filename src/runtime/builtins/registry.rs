//! Token 注册表

use std::collections::HashMap;
use super::token::{BuiltinToken, TOKEN_REGISTRY};

/// Token 注册表
/// 维护函数名到 Token ID 的映射
#[allow(dead_code)]
pub struct TokenRegistry {
    /// 函数名到 Token 的映射（小写存储，不区分大小写）
    map: HashMap<String, BuiltinToken>,
}

impl TokenRegistry {
    /// 创建新的 Token 注册表
    pub fn new() -> Self {
        let map = TOKEN_REGISTRY
            .iter()
            .map(|(name, token)| (name.to_string(), *token))
            .collect();
        Self { map }
    }

    /// 查找函数名对应的 Token
    pub fn lookup(&self, name: &str) -> Option<BuiltinToken> {
        self.map.get(&name.to_lowercase()).copied()
    }

    /// 检查是否为内置函数
    #[allow(dead_code)]
    pub fn is_builtin(&self, name: &str) -> bool {
        self.map.contains_key(&name.to_lowercase())
    }
}

impl Default for TokenRegistry {
    fn default() -> Self {
        Self::new()
    }
}
