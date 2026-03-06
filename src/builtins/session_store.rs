//! Session 存储后端 trait 定义
//!
//! 支持多种存储后端：内存、JSON文件、Redis（预留）

use super::session::SessionData;
use std::collections::HashMap;

/// Session 存储后端 trait
pub trait SessionStore: Send + Sync {
    /// 获取 Session
    fn get(&self, session_id: &str) -> Option<SessionData>;

    /// 保存 Session
    fn set(&mut self, session_id: &str, data: SessionData);

    /// 删除 Session
    fn delete(&mut self, session_id: &str);

    /// 获取所有 Session ID
    fn keys(&self) -> Vec<String>;

    /// 清理过期 Session（返回清理数量）
    fn cleanup(&mut self, now: u64) -> usize;
}

/// 内存存储后端
pub struct MemoryStore {
    data: HashMap<String, SessionData>,
}

impl MemoryStore {
    pub fn new() -> Self {
        MemoryStore {
            data: HashMap::new(),
        }
    }
}

impl SessionStore for MemoryStore {
    fn get(&self, session_id: &str) -> Option<SessionData> {
        self.data.get(session_id).cloned()
    }

    fn set(&mut self, session_id: &str, data: SessionData) {
        self.data.insert(session_id.to_string(), data);
    }

    fn delete(&mut self, session_id: &str) {
        self.data.remove(session_id);
    }

    fn keys(&self) -> Vec<String> {
        self.data.keys().cloned().collect()
    }

    fn cleanup(&mut self, now: u64) -> usize {
        let expired: Vec<String> = self
            .data
            .iter()
            .filter(|(_, v)| {
                let timeout_seconds = v.timeout as u64 * 60;
                now > v.last_accessed + timeout_seconds
            })
            .map(|(k, _)| k.clone())
            .collect();

        let count = expired.len();
        for key in expired {
            self.data.remove(&key);
        }
        count
    }
}

/// JSON 文件存储后端
pub struct JsonFileStore {
    base_path: std::path::PathBuf,
}

impl JsonFileStore {
    pub fn new(base_path: std::path::PathBuf) -> Self {
        // 确保目录存在
        if !base_path.exists() {
            std::fs::create_dir_all(&base_path).ok();
        }
        JsonFileStore { base_path }
    }

    fn file_path(&self, session_id: &str) -> std::path::PathBuf {
        self.base_path.join(format!("{}.json", session_id))
    }
}

impl SessionStore for JsonFileStore {
    fn get(&self, session_id: &str) -> Option<SessionData> {
        let path = self.file_path(session_id);
        let content = std::fs::read_to_string(&path).ok()?;
        serde_json::from_str(&content).ok()
    }

    fn set(&mut self, session_id: &str, data: SessionData) {
        let path = self.file_path(session_id);
        if let Ok(json) = serde_json::to_string_pretty(&data) {
            std::fs::write(&path, json).ok();
        }
    }

    fn delete(&mut self, session_id: &str) {
        let path = self.file_path(session_id);
        std::fs::remove_file(&path).ok();
    }

    fn keys(&self) -> Vec<String> {
        let mut keys = Vec::new();
        if let Ok(entries) = std::fs::read_dir(&self.base_path) {
            for entry in entries.flatten() {
                if let Some(name) = entry.file_name().to_str() {
                    if name.ends_with(".json") {
                        let key = name.trim_end_matches(".json").to_string();
                        keys.push(key);
                    }
                }
            }
        }
        keys
    }

    fn cleanup(&mut self, now: u64) -> usize {
        let mut count = 0;
        let keys = self.keys();
        for key in keys {
            if let Some(data) = self.get(&key) {
                let timeout_seconds = data.timeout as u64 * 60;
                if now > data.last_accessed + timeout_seconds {
                    self.delete(&key);
                    count += 1;
                }
            }
        }
        count
    }
}

/// Redis 存储后端（预留接口）
#[cfg(feature = "redis")]
pub struct RedisStore {
    client: redis::Client,
    prefix: String,
}

#[cfg(feature = "redis")]
impl RedisStore {
    pub fn new(redis_url: &str, prefix: &str) -> Result<Self, Box<dyn std::error::Error>> {
        let client = redis::Client::open(redis_url)?;
        Ok(RedisStore {
            client,
            prefix: prefix.to_string(),
        })
    }

    fn key(&self, session_id: &str) -> String {
        format!("{}{}", self.prefix, session_id)
    }
}

#[cfg(feature = "redis")]
impl SessionStore for RedisStore {
    fn get(&self, session_id: &str) -> Option<SessionData> {
        let mut conn = self.client.get_connection().ok()?;
        let json: String = redis::Cmd::get(&self.key(session_id))
            .query(&mut conn)
            .ok()?;
        serde_json::from_str(&json).ok()
    }

    fn set(&mut self, session_id: &str, data: SessionData) {
        if let Ok(json) = serde_json::to_string(&data) {
            let ttl = data.timeout as usize * 60;
            if let Ok(mut conn) = self.client.get_connection() {
                let key = self.key(session_id);
                let _: Result<(), _> = redis::Cmd::set_ex(&key, json, ttl).query(&mut conn);
            }
        }
    }

    fn delete(&mut self, session_id: &str) {
        if let Ok(mut conn) = self.client.get_connection() {
            let _: Result<(), _> = redis::Cmd::del(&self.key(session_id)).query(&mut conn);
        }
    }

    fn keys(&self) -> Vec<String> {
        // Redis keys 扫描实现较复杂，简化返回空
        Vec::new()
    }

    fn cleanup(&mut self, _now: u64) -> usize {
        // Redis TTL 自动处理过期，无需手动清理
        0
    }
}

/// 根据配置创建存储后端
pub fn create_store(storage_type: &str, runtime_dir: &std::path::Path) -> Box<dyn SessionStore> {
    match storage_type {
        "memory" => {
            eprintln!("[Session] 使用内存 Session 存储");
            Box::new(MemoryStore::new())
        }
        "json" => {
            let path = runtime_dir.join("sessions");
            eprintln!("[Session] 使用 JSON 文件 Session 存储: {:?}", path);
            Box::new(JsonFileStore::new(path))
        }
        "redis" => {
            eprintln!("[Session] Redis 存储尚未实现，回退到内存存储");
            Box::new(MemoryStore::new())
        }
        _ => {
            eprintln!("[Session] 未知的存储类型: {}，使用内存存储", storage_type);
            Box::new(MemoryStore::new())
        }
    }
}
