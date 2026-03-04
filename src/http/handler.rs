//! HTTP 处理器

use axum::{
    body::Body,
    http::{StatusCode, Uri},
    response::{Html, IntoResponse, Response},
};
use std::path::Path;
use std::path::PathBuf;
use tokio::fs;

use super::path_resolver::PathResolver;
use super::state::AppState;

/// 格式化详细错误信息
fn format_error(
    location: &str,
    line: u32,
    uri: &str,
    file_path: &Path,
    error_type: &str,
    error_msg: &str,
) -> String {
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
        location, line, error_type, uri, file_path.display(), error_msg
    )
}

/// 处理 ASP 请求（使用简化引擎）
pub async fn handle_asp(uri: Uri, state: AppState) -> impl IntoResponse {
    let uri_str = uri.path();

    // 使用路径解析器安全解析路径
    let resolver = PathResolver::new(
        state.config.home_dir.clone(),
        state.config.allow_parent_paths,
    );

    let file_path = match resolver.resolve(uri_str) {
        Ok(path) => path,
        Err(e) => {
            return Html(format_error(
                "handler.rs",
                22, // resolver.resolve 调用行号
                uri_str,
                Path::new("(path resolution failed)"),
                "Path Resolution Error",
                &e.to_string(),
            ))
            .into_response();
        }
    };

    // 检查文件是否存在
    if !file_path.exists() {
        return Html(format_error(
            "handler.rs",
            31,
            uri_str,
            &file_path,
            "File Not Found",
            "The requested file does not exist.",
        ))
        .into_response();
    }

    // 检查是否是目录
    if file_path.is_dir() {
        return Html(format_error(
            "handler.rs",
            52,
            uri_str,
            &file_path,
            "Is a Directory",
            &format!(
                "The requested path is a directory, not a file.\n\nHint: Try accessing {}/",
                uri_str.trim_end_matches('/')
            ),
        ))
        .into_response();
    }

    // 读取文件内容
    match fs::read_to_string(&file_path).await {
        Ok(content) => {
            // 使用简化的 ASP 引擎执行
            let mut engine = crate::asp::Engine::new().with_debug(state.config.debug);
            match engine.execute(&content) {
                Ok(output) => Html(output).into_response(),
                Err(e) => Html(format_error(
                    "simple_engine.rs",
                    0, // 引擎内部错误
                    uri_str,
                    &file_path,
                    "ASP Execution Error",
                    &e.to_string(),
                ))
                .into_response(),
            }
        }
        Err(e) => Html(format_error(
            "handler.rs",
            36,
            uri_str,
            &file_path,
            "File Read Error",
            &e.to_string(),
        ))
        .into_response(),
    }
}

/// 处理静态文件请求
pub async fn handle_static(uri: Uri, state: AppState) -> impl IntoResponse {
    let uri_str = uri.path();

    // 使用路径解析器安全解析路径
    let resolver = PathResolver::new(
        state.config.home_dir.clone(),
        state.config.allow_parent_paths,
    );

    let file_path = match resolver.resolve(uri_str) {
        Ok(path) => path,
        Err(e) => {
            return Response::builder()
                .status(StatusCode::FORBIDDEN)
                .body(Body::from(format!("Forbidden: {}", e)))
                .unwrap();
        }
    };

    // 检查是否是目录
    if file_path.is_dir() {
        // 尝试返回索引文件
        let index_path = file_path.join(&state.config.index_file);
        if index_path.exists() {
            return handle_asp(uri, state).await.into_response();
        }

        // 显示目录列表
        if state.config.directory_listing {
            return generate_directory_listing(&file_path, uri.path()).await;
        }

        return Response::builder()
            .status(StatusCode::FORBIDDEN)
            .body(Body::from("Directory listing is disabled"))
            .unwrap();
    }

    // 检查文件是否存在
    if !file_path.exists() {
        return Response::builder()
            .status(StatusCode::NOT_FOUND)
            .body(Body::from(format_error(
                "handler.rs",
                87,
                uri_str,
                &file_path,
                "File Not Found",
                "The requested file does not exist.",
            )))
            .unwrap();
    }

    // 读取并返回文件
    match fs::read(&file_path).await {
        Ok(content) => {
            let mime = mime_guess::from_path(&file_path)
                .first_or_octet_stream()
                .to_string();
            Response::builder()
                .status(StatusCode::OK)
                .header("Content-Type", mime)
                .body(Body::from(content))
                .unwrap()
        }
        Err(e) => {
            let error_html = format_error(
                "handler.rs",
                95,
                uri_str,
                &file_path,
                "File Read Error",
                &e.to_string(),
            );
            Response::builder()
                .status(StatusCode::INTERNAL_SERVER_ERROR)
                .body(Body::from(error_html))
                .unwrap()
        }
    }
}

/// 生成目录列表
pub async fn generate_directory_listing(dir: &PathBuf, url_path: &str) -> Response {
    let mut html = String::new();
    html.push_str("<!DOCTYPE HTML>\n<html lang=\"en\">\n<head>\n");
    html.push_str("<meta charset=\"utf-8\">\n");
    html.push_str(&format!(
        "<title>Directory listing for {}</title>\n",
        url_path
    ));
    html.push_str("</head>\n<body>\n");
    html.push_str(&format!("<h1>Directory listing for {}</h1>\n", url_path));
    html.push_str("<hr>\n<ul>\n");

    if let Ok(mut entries) = fs::read_dir(dir).await {
        let mut entry_list = Vec::new();
        while let Ok(Some(entry)) = entries.next_entry().await {
            entry_list.push(entry);
        }
        entry_list.sort_by_key(|e| e.file_name());

        for entry in entry_list {
            if let Ok(name) = entry.file_name().into_string() {
                let is_dir = entry.file_type().await.map(|t| t.is_dir()).unwrap_or(false);
                let display_name = if is_dir {
                    format!("{}/", name)
                } else {
                    name.clone()
                };

                // 构建正确的 URL 路径
                let href = if url_path == "/" || url_path.is_empty() {
                    format!("/{}", name)
                } else {
                    // 移除 url_path 末尾的 /（如果有），然后拼接
                    let base = url_path.trim_end_matches('/');
                    format!("{}/{}", base, name)
                };

                html.push_str(&format!(
                    "<li><a href=\"{}\">{}</a></li>\n",
                    href, display_name
                ));
            }
        }
    }

    html.push_str("</ul>\n<hr>\n</body>\n</html>");
    Html(html).into_response()
}
