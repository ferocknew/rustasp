//! 错误页面生成模块

use axum::{
    body::Body,
    http::{Response, StatusCode},
};
use tokio::fs;

use super::state::AppState;

/// 错误类型枚举
#[derive(Debug)]
pub enum ErrorKind {
    PathResolution,
    FileNotFound,
    DirectoryListingDisabled,
    FileRead,
    AspExecution,
}

/// 错误信息
pub struct ErrorInfo {
    pub location: &'static str,
    pub line: u32,
    pub uri: String,
    pub file_path: String,
    pub kind: ErrorKind,
    pub message: String,
}

impl ErrorInfo {
    pub fn new(
        location: &'static str,
        line: u32,
        uri: impl Into<String>,
        file_path: impl Into<String>,
        kind: ErrorKind,
        message: impl Into<String>,
    ) -> Self {
        Self {
            location,
            line,
            uri: uri.into(),
            file_path: file_path.into(),
            kind,
            message: message.into(),
        }
    }

    /// 获取错误类型标题
    fn title(&self) -> &'static str {
        match self.kind {
            ErrorKind::PathResolution => "Path Resolution Error",
            ErrorKind::FileNotFound => "File Not Found",
            ErrorKind::DirectoryListingDisabled => "Directory Listing Disabled",
            ErrorKind::FileRead => "File Read Error",
            ErrorKind::AspExecution => "ASP Execution Error",
        }
    }

    /// 获取 HTTP 状态码
    pub fn status_code(&self) -> StatusCode {
        match self.kind {
            ErrorKind::PathResolution => StatusCode::FORBIDDEN,
            ErrorKind::FileNotFound => StatusCode::NOT_FOUND,
            ErrorKind::DirectoryListingDisabled => StatusCode::FORBIDDEN,
            ErrorKind::FileRead => StatusCode::INTERNAL_SERVER_ERROR,
            ErrorKind::AspExecution => StatusCode::INTERNAL_SERVER_ERROR,
        }
    }
}

/// 错误页面生成器
pub struct ErrorPageGenerator {
    detailed_error: bool,
    custom_error_page: Option<String>,
}

impl ErrorPageGenerator {
    /// 从 AppState 创建，并加载自定义错误页面
    pub async fn from_state(state: &AppState) -> Self {
        let custom_page = load_custom_error_page(state).await;
        Self {
            detailed_error: state.config.detailed_error,
            custom_error_page: custom_page,
        }
    }

    /// 生成错误响应
    pub fn generate(&self, error: &ErrorInfo) -> Response<Body> {
        let html = if self.detailed_error {
            format_detailed_error(error)
        } else {
            self.get_simple_error()
        };

        Response::builder()
            .status(error.status_code())
            .body(Body::from(html))
            .unwrap()
    }

    fn get_simple_error(&self) -> String {
        self.custom_error_page
            .clone()
            .unwrap_or_else(format_simple_error)
    }
}

/// 格式化简单错误信息
fn format_simple_error() -> String {
    r#"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>500 Internal Server Error</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 40px; text-align: center; }
        h1 { color: #d32f2f; }
    </style>
</head>
<body>
    <h1>500 Internal Server Error</h1>
    <p>An error occurred while processing your request.</p>
</body>
</html>
"#
    .to_string()
}

/// 格式化详细错误信息
fn format_detailed_error(error: &ErrorInfo) -> String {
    format!(
        r#"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Error</title>
    <style>
        body {{ font-family: monospace; padding: 20px; background: #f5f5f5; }}
        .error-box {{ background: white; border-left: 4px solid #d32f2f; padding: 15px; margin: 10px 0; }}
        .location {{ color: #666; font-size: 12px; }}
        .error-type {{ color: #d32f2f; font-weight: bold; margin: 10px 0; }}
        .details {{ margin: 10px 0; }}
        .label {{ color: #1976d2; font-weight: bold; }}
        pre {{ background: #f5f5f5; padding: 10px; overflow-x: auto; }}
    </style>
</head>
<body>
    <div class="error-box">
        <div class="location">📍 {}:{}</div>
        <h1>Error</h1>
        <div class="error-type">❌ {}</div>
        <div class="details">
            <div><span class="label">Request URI:</span> <pre>{}</pre></div>
            <div><span class="label">File Path:</span> <pre>{}</pre></div>
            <div><span class="label">Details:</span> <pre>{}</pre></div>
        </div>
    </div>
</body>
</html>
"#,
        error.location,
        error.line,
        error.title(),
        error.uri,
        error.file_path,
        error.message
    )
}

/// 读取自定义错误页面
async fn load_custom_error_page(state: &AppState) -> Option<String> {
    let error_page_path = state.config.error_page.as_ref()?;
    let full_path = state.config.home_dir.join(error_page_path);
    fs::read_to_string(&full_path).await.ok()
}
