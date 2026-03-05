//! ASP 引擎模块

mod engine;
mod include;
mod segmenter;

pub use engine::{Engine, ExecutionResult};
pub use include::preprocess;
