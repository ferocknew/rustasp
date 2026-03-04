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
    /// 源代码（可选，用于 ASP 执行错误）
    pub source_code: Option<String>,
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
            source_code: None,
        }
    }

    /// 设置源代码
    pub fn with_source_code(mut self, code: impl Into<String>) -> Self {
        self.source_code = Some(code.into());
        self
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
    // 格式化代码块（带行号）
    let code_section = if let Some(ref code) = error.source_code {
        let error_line = extract_error_line(&error.message);
        format_code_with_lines(code, error_line)
    } else {
        String::new()
    };

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
        pre {{ background: #f5f5f5; padding: 10px; overflow-x: auto; margin: 5px 0; }}
        .code-container {{ margin: 10px 0; }}
        .code-block {{ background: #282c34; color: #abb2bf; padding: 12px; overflow-x: auto; border-radius: 4px; }}
        .code-block .line {{ display: flex; min-height: 1.4em; }}
        .code-block .line.error-line {{ background: rgba(229, 115, 115, 0.2); border-left: 3px solid #e57373; }}
        .code-block .line-num {{ color: #636d83; min-width: 50px; padding-right: 15px; text-align: right; user-select: none; }}
        .code-block .line-content {{ white-space: pre; flex: 1; }}
        .error-code {{ background: #fff3e0; border-left: 3px solid #ff9800; padding: 10px; margin: 5px 0; white-space: pre-wrap; word-break: break-all; }}
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
            {}
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
        html_escape(&error.message),
        code_section
    )
}

/// 从错误消息中提取行号
fn extract_error_line(message: &str) -> Option<usize> {
    // 尝试匹配 "at line X" 或 "行 X" 等模式
    let lower = message.to_lowercase();

    // 尝试找 "at line X" 模式
    if let Some(pos) = lower.find("at line") {
        let rest = &message[pos + 7..];
        let rest = rest.trim_start();
        if let Some(num_end) = rest.find(|c: char| !c.is_ascii_digit()) {
            if let Ok(line) = rest[..num_end].parse::<usize>() {
                return Some(line);
            }
        } else if let Ok(line) = rest.parse::<usize>() {
            return Some(line);
        }
    }

    // 尝试找 "line X" 模式
    if let Some(pos) = lower.find("line") {
        let rest = &message[pos + 4..];
        let rest = rest.trim_start();
        if let Some(num_end) = rest.find(|c: char| !c.is_ascii_digit()) {
            if let Ok(line) = rest[..num_end].parse::<usize>() {
                return Some(line);
            }
        }
    }

    None
}

/// 格式化代码并添加行号，可选高亮错误行
fn format_code_with_lines(code: &str, error_line: Option<usize>) -> String {
    let lines: Vec<&str> = code.lines().collect();
    let total_lines = lines.len();

    // 确定显示范围（错误行前后各显示上下文）
    let (start, end) = if let Some(err_line) = error_line {
        let err_idx = err_line.saturating_sub(1); // 转为 0-indexed
        let context = 2; // 前后各 2 行
        let start = err_idx.saturating_sub(context);
        let end = (err_idx + context + 1).min(total_lines);
        (start, end)
    } else {
        // 没有错误行信息，显示前 20 行或全部
        (0, total_lines.min(20))
    };

    let mut result = String::new();
    result.push_str(r#"<div class="code-container"><span class="label">Source Code:</span><div class="code-block">"#);

    for (idx, line) in lines.iter().enumerate().skip(start).take(end - start) {
        let line_num = idx + 1;
        let is_error = error_line == Some(line_num);
        let error_class = if is_error { " error-line" } else { "" };

        result.push_str(&format!(
            r#"<div class="line{}"><span class="line-num">{}</span><span class="line-content">{}</span></div>"#,
            error_class,
            line_num,
            html_escape(line)
        ));
    }

    // 如果截断了代码，显示省略提示
    if end < total_lines {
        result.push_str(&format!(
            r#"<div class="line"><span class="line-num">...</span><span class="line-content">/* {} more lines */</span></div>"#,
            total_lines - end
        ));
    }

    result.push_str("</div></div>");
    result
}

/// HTML 转义
fn html_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// 读取自定义错误页面
async fn load_custom_error_page(state: &AppState) -> Option<String> {
    let error_page_path = state.config.error_page.as_ref()?;
    let full_path = state.config.home_dir.join(error_page_path);
    fs::read_to_string(&full_path).await.ok()
}
