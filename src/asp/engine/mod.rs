//! ASP 完整引擎
//!
//! 使用 Parser + Runtime 执行 ASP 代码
//! 支持跨代码块的语句（如 If...Else...End If 分散在多个代码块中）

mod error;
mod executor;

pub use executor::Engine;

use std::collections::HashMap;
use vbscript::runtime::objects::Session;
use vbscript::runtime::Value;
use vbscript::Response;

/// ASP 执行结果
pub struct ExecutionResult {
    /// 输出内容
    pub output: String,
    /// Response 对象（包含状态码、ContentType、Headers 等）
    pub response: Response,
}

/// 将 Session 转换为 HashMap
#[allow(dead_code)]
pub(super) fn session_to_map(session: &Session) -> HashMap<String, Value> {
    let mut map = HashMap::new();
    // 存储 Session ID
    map.insert(
        "sessionid".to_string(),
        Value::String(session.session_id().to_string()),
    );
    map.insert(
        "timeout".to_string(),
        Value::Number(session.timeout() as f64),
    );
    map
}
