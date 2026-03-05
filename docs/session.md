# Rust ASP 引擎 Session 实现说明（文件存储）

## 1. 设计目标

为 ASP / VBScript 引擎实现 "Session" 对象，使脚本可以读写会话数据：

```asp
Session("user") = "Tom"
Response.Write Session("user")
```

设计目标：

- 实现简单
- 无数据库依赖
- 方便调试
- 支持多用户访问

Session 使用 Cookie + 文件存储 实现。

---

## 2. 工作流程

HTTP 请求处理流程：

```
HTTP Request
     │
     ▼
读取 Cookie: ASPSESSIONID
     │
     ├─ 如果不存在
     │     生成新的 SessionID
     │
     ▼
加载 Session 文件
     │
     ▼
脚本运行 (读写 Session)
     │
     ▼
保存 Session 文件
     │
     ▼
HTTP Response + Set-Cookie
```

---

## 3. SessionID

SessionID 用于唯一标识用户会话。

建议使用 UUID 生成：

```rust
use uuid::Uuid;

let session_id = Uuid::new_v4().to_string();
```

示例：

```
3e9c4a6f-9d63-4b1a-b72f-3c8c8c7b0d9e
```

浏览器 Cookie：

```
Set-Cookie: ASPSESSIONID=3e9c4a6f-9d63-4b1a-b72f-3c8c8c7b0d9e
```

---

## 4. Session 文件结构

Session 数据存储在服务器文件中。

目录结构：

```
runtime/
   sessions/
```

Session 文件：

```
runtime/sessions/<session_id>.json
```

示例：

```
runtime/sessions/3e9c4a6f.json
```

---

## 5. Session 数据格式

Session 文件使用 JSON 保存：

```json
{
  "user": "Tom",
  "user_id": 123
}
```

对应 Rust 结构：

```rust
use std::collections::HashMap;

pub struct Session {
    pub id: String,
    pub data: HashMap<String, Value>,
}
```

---

## 6. 读取 Session

处理请求时：

1. 从 Cookie 获取 SessionID
2. 查找对应 Session 文件
3. 如果存在则加载
4. 如果不存在则创建新的 Session

示例：

```rust
fn load_session(id: &str) -> Option<Session> {
    let path = format!("runtime/sessions/{}.json", id);

    if !std::path::Path::new(&path).exists() {
        return None;
    }

    let text = std::fs::read_to_string(path).ok()?;

    serde_json::from_str(&text).ok()
}
```

---

## 7. 保存 Session

请求结束时，将 Session 写回文件：

```rust
fn save_session(session: &Session) {
    let path = format!("runtime/sessions/{}.json", session.id);

    let text = serde_json::to_string(&session.data).unwrap();

    std::fs::write(path, text).unwrap();
}
```

---

## 8. VBScript 接口

VBScript 中：

```asp
Session("user") = "Tom"
```

解释器调用：

```rust
session.set("user", value);
```

读取：

```rust
session.get("user");
```

示例实现：

```rust
impl Session {
    pub fn get(&self, key: &str) -> Option<&Value> {
        self.data.get(key)
    }

    pub fn set(&mut self, key: String, value: Value) {
        self.data.insert(key, value);
    }
}
```

---

## 9. Session 过期

Session 需要设置过期时间。

经典 ASP 默认：

```
20 分钟
```

可以在 Session 文件中记录时间：

```json
{
  "data": {...},
  "last_access": 1710000000
}
```

服务器定期清理过期 Session 文件。

---

## 10. 并发问题

如果服务器是多线程处理请求，写 Session 文件时需要加锁。

简单方案：

- 每个 Session 一个文件
- 写入时使用文件锁

避免多个请求同时写入导致数据损坏。

---

## 11. 总结

Session 实现结构：

```
Browser Cookie
       │
       ▼
SessionID
       │
       ▼
Session 文件
       │
       ▼
HashMap<String, Value>
```

特点：

- 实现简单
- 易于维护
- 调试方便
- 适合 ASP 引擎初期版本
