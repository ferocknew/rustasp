//! HTTP 服务模块
//!
//! 负责：路由、文件加载、构建 Request 上下文、返回 HTTP 响应

mod handler;
mod router;
pub mod state;

pub use router::create_router;
pub use state::AppState;
