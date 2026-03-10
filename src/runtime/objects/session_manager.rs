//! Session 管理器
//!
//! 负责 Session 的加载、保存、清理和 Cookie 集成

use super::session::{Session, SessionData};
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

/// Session 管理器
pub struct SessionManager {
    /// Session 存储目录
    session_dir: PathBuf,
    /// 内存缓存（可选，用于性能优化）
    cache: HashMap<String, Session>,
}

impl SessionManager {
    /// 创建新的 Session 管理器
    pub fn new<P: AsRef<Path>>(runtime_dir: P) -> Result<Self, String> {
        let session_dir = runtime_dir.as_ref().join("sessions");

        // 创建 Session 目录（如果不存在）
        if !session_dir.exists() {
            fs::create_dir_all(&session_dir)
                .map_err(|e| format!("无法创建 Session 目录: {}", e))?;
        }

        Ok(SessionManager {
            session_dir,
            cache: HashMap::new(),
        })
    }

    /// 生成新的 Session ID
    pub fn generate_session_id() -> String {
        // 使用 UUID v4 格式
        let uuid = uuid::Uuid::new_v4();
        uuid.to_string().replace("-", "")
    }

    /// 创建新的 Session
    pub fn create_session(&mut self, session_id: String, timeout: u32) -> Result<Session, String> {
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();

        let session_data = SessionData {
            session_id: session_id.clone(),
            timeout,
            created_at: now,
            last_accessed: now,
            data: HashMap::new(),
        };

        // 保存到文件
        self.save_session_data(&session_data)?;

        // 创建 Session 对象
        let session = Session::new(session_id.clone());
        self.cache.insert(session_id, session.clone());

        Ok(session)
    }

    /// 加载 Session
    pub fn load_session(&mut self, session_id: &str) -> Result<Option<Session>, String> {
        // 检查缓存
        if let Some(session) = self.cache.get(session_id) {
            return Ok(Some(session.clone()));
        }

        // 从文件加载
        let session_data = match self.load_session_data(session_id)? {
            Some(data) => data,
            None => return Ok(None),
        };

        // 检查是否过期
        if self.is_expired(&session_data) {
            // 删除过期 Session
            self.delete_session(session_id)?;
            return Ok(None);
        }

        // 更新最后访问时间
        let mut updated_data = session_data.clone();
        updated_data.last_accessed = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        self.save_session_data(&updated_data)?;

        // 从 SessionData 创建 Session 对象（恢复数据）
        let session = Session::from_session_data(updated_data);
        self.cache.insert(session_id.to_string(), session.clone());

        Ok(Some(session))
    }

    /// 保存 Session
    pub fn save_session(&mut self, session: &Session) -> Result<(), String> {
        // 转换为 SessionData
        let session_data = session.to_session_data()?;

        // 保存到文件
        self.save_session_data(&session_data)?;

        // 更新缓存
        self.cache
            .insert(session.session_id().to_string(), session.clone());

        Ok(())
    }

    /// 保存 Session 数据到文件
    pub fn save_session_data(&self, session_data: &SessionData) -> Result<(), String> {
        let file_path = self
            .session_dir
            .join(format!("{}.json", session_data.session_id));

        let json = serde_json::to_string_pretty(session_data)
            .map_err(|e| format!("无法序列化 Session 数据: {}", e))?;

        fs::write(&file_path, json).map_err(|e| format!("无法写入 Session 文件: {}", e))?;

        Ok(())
    }

    /// 从文件加载 Session 数据
    fn load_session_data(&self, session_id: &str) -> Result<Option<SessionData>, String> {
        let file_path = self.session_dir.join(format!("{}.json", session_id));

        if !file_path.exists() {
            return Ok(None);
        }

        let json =
            fs::read_to_string(&file_path).map_err(|e| format!("无法读取 Session 文件: {}", e))?;

        let session_data: SessionData =
            serde_json::from_str(&json).map_err(|e| format!("无法解析 Session 数据: {}", e))?;

        Ok(Some(session_data))
    }

    /// 删除 Session
    fn delete_session(&self, session_id: &str) -> Result<(), String> {
        let file_path = self.session_dir.join(format!("{}.json", session_id));

        if file_path.exists() {
            fs::remove_file(&file_path).map_err(|e| format!("无法删除 Session 文件: {}", e))?;
        }

        Ok(())
    }

    /// 检查 Session 是否过期
    fn is_expired(&self, session_data: &SessionData) -> bool {
        let now = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();

        // timeout 单位为秒
        now > session_data.last_accessed + session_data.timeout as u64
    }

    /// 清理所有过期 Session
    pub fn cleanup_expired_sessions(&self) -> Result<usize, String> {
        let mut cleaned_count = 0;

        let entries =
            fs::read_dir(&self.session_dir).map_err(|e| format!("无法读取 Session 目录: {}", e))?;

        for entry in entries {
            let entry = entry.map_err(|e| format!("无法读取目录项: {}", e))?;
            let path = entry.path();

            // 只处理 .json 文件
            if path.extension().and_then(|s| s.to_str()) != Some("json") {
                continue;
            }

            // 尝试加载并检查是否过期
            let file_name = path.file_stem().and_then(|s| s.to_str()).unwrap_or("");

            if let Ok(Some(session_data)) = self.load_session_data(file_name) {
                if self.is_expired(&session_data) {
                    self.delete_session(file_name)?;
                    cleaned_count += 1;
                }
            }
        }

        Ok(cleaned_count)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn test_session_manager_creation() {
        let temp_dir = TempDir::new().unwrap();
        let _manager = SessionManager::new(temp_dir.path()).unwrap();

        // 验证 sessions 目录被创建
        let sessions_dir = temp_dir.path().join("sessions");
        assert!(sessions_dir.exists());
    }

    #[test]
    fn test_create_and_load_session() {
        let temp_dir = TempDir::new().unwrap();
        let mut manager = SessionManager::new(temp_dir.path()).unwrap();

        // 创建 Session
        let session_id = SessionManager::generate_session_id();
        let _session = manager.create_session(session_id.clone(), 20).unwrap();

        // 重新加载 Session
        let loaded = manager.load_session(&session_id).unwrap();
        assert!(loaded.is_some());
    }

    #[test]
    fn test_session_expiration() {
        let temp_dir = TempDir::new().unwrap();
        let mut manager = SessionManager::new(temp_dir.path()).unwrap();

        // 创建过期时间的 Session（1秒）
        let session_id = SessionManager::generate_session_id();
        manager.create_session(session_id.clone(), 1).unwrap();

        // 清除缓存，确保从文件加载
        manager.cache.clear();

        // 等待 Session 过期
        std::thread::sleep(std::time::Duration::from_secs(2));

        // 尝试加载过期 Session
        let loaded = manager.load_session(&session_id).unwrap();
        assert!(loaded.is_none()); // 应该返回 None（已过期）
    }

    #[test]
    fn test_cleanup_expired_sessions() {
        let temp_dir = TempDir::new().unwrap();
        let mut manager = SessionManager::new(temp_dir.path()).unwrap();

        // 创建两个 Session，一个快速过期
        let session1 = SessionManager::generate_session_id();
        let session2 = SessionManager::generate_session_id();

        manager.create_session(session1.clone(), 1).unwrap(); // 1秒过期
        manager.create_session(session2.clone(), 60).unwrap(); // 60秒过期

        // 等待第一个 Session 过期
        std::thread::sleep(std::time::Duration::from_secs(2));

        // 清理过期 Session
        let cleaned = manager.cleanup_expired_sessions().unwrap();
        assert_eq!(cleaned, 1); // 应该清理掉1个
    }
}
