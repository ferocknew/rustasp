# Session 存储架构说明

## 概述

本项目实现了**可插拔的 Session 存储架构**，支持多种存储后端：

1. **内存存储 (memory)** - 纯内存，重启后数据丢失，性能最高
2. **JSON文件存储 (json)** - 文件持久化，适合单机部署
3. **Redis存储 (redis)** - 分布式缓存，适合多实例部署（预留接口）

## 配置方式

通过 `.env` 文件中的环境变量切换存储模式：

```bash
# Session 存储模式 (memory/json/redis)
SESSION_STORAGE=memory

# Session 超时时间（分钟）
SESSION_TIMEOUT=20

# Session 存储目录（仅 json 模式使用）
SESSION_DIR=./runtime/sessions

# Redis 连接地址（预留，仅 redis 模式使用）
# REDIS_URL=redis://127.0.0.1:6379
# REDIS_KEY_PREFIX=vbscript:session:
```

## 使用示例

### 1. 内存存储模式（开发/测试）

```bash
# .env
SESSION_STORAGE=memory
SESSION_TIMEOUT=20
```

特点：
- 数据保存在内存中，访问速度最快
- 应用重启后 Session 数据丢失
- 适合开发和测试环境

### 2. JSON 文件存储模式（生产单实例）

```bash
# .env
SESSION_STORAGE=json
SESSION_TIMEOUT=60
SESSION_DIR=./runtime/sessions
```

特点：
- Session 数据持久化到 JSON 文件
- 应用重启后数据保留
- 适合单机生产环境

### 3. Redis 存储模式（生产多实例）- 预留

```bash
# .env
SESSION_STORAGE=redis
SESSION_TIMEOUT=60
REDIS_URL=redis://127.0.0.1:6379
REDIS_KEY_PREFIX=vbscript:session:
```

特点：
- Session 数据存储在 Redis
- 支持多实例共享 Session
- 适合分布式部署

## 架构设计

### 核心 Trait

```rust
/// Session 存储后端 trait
pub trait SessionStore: Send + Sync {
    /// 获取 Session
    fn get(&self, session_id: &str) -> Option<SessionData>;

    /// 保存 Session
    fn set(&mut self, session_id: &str, data: SessionData);

    /// 删除 Session
    fn delete(&mut self, session_id: &str);

    /// 清理过期 Session
    fn cleanup(&mut self, now: u64) -> usize;
}
```

### 存储后端实现

1. **MemoryStore** - 基于 `HashMap<String, SessionData>`
2. **JsonFileStore** - 基于文件系统 + JSON 序列化
3. **RedisStore** - 基于 Redis（预留接口）

### 工厂函数

```rust
/// 根据配置创建存储后端
pub fn create_store(
    storage_type: &str,
    config: &Config
) -> Box<dyn SessionStore> {
    match storage_type {
        "memory" => Box::new(MemoryStore::new()),
        "json" => Box::new(JsonFileStore::new(path)),
        "redis" => Box::new(RedisStore::new(...)),
        _ => Box::new(MemoryStore::new()),
    }
}
```

## 扩展指南

### 添加新的存储后端

1. 实现 `SessionStore` trait
2. 在 `create_store` 中添加分支
3. 在 `.env.example` 中添加配置项

示例：

```rust
// 1. 实现存储后端
pub struct MyCustomStore { ... }

impl SessionStore for MyCustomStore {
    fn get(&self, session_id: &str) -> Option<SessionData> { ... }
    fn set(&mut self, session_id: &str, data: SessionData) { ... }
    fn delete(&mut self, session_id: &str) { ... }
    fn cleanup(&mut self, now: u64) -> usize { ... }
}

// 2. 在 create_store 中添加
"custom" => Box::new(MyCustomStore::new()),
```

## 注意事项

1. **内存模式**：适合开发和测试，生产环境建议用 json 或 redis
2. **JSON 模式**：确保 `SESSION_DIR` 目录可写
3. **Redis 模式**：需要启用 `redis` feature 编译
4. **切换存储**：切换后原有 Session 数据会丢失（不同存储之间不共享）