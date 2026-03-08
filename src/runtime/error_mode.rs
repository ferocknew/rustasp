//! 错误模式

/// 错误处理模式
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorMode {
    /// 默认模式：遇到错误停止执行
    Stop,
    /// Resume Next 模式：遇到错误继续执行下一条语句
    ResumeNext,
}

impl Default for ErrorMode {
    fn default() -> Self {
        ErrorMode::Stop
    }
}
