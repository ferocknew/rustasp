//! ASP 引擎模块
//!
//! 负责：分割 HTML 与 `<% %>`、执行脚本片段、拼接输出、管理单次请求上下文

mod engine;
mod segmenter;

pub use engine::Engine;
pub use segmenter::Segment;
