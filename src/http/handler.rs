//! HTTP 处理器

use axum::{
    body::Body,
    http::{Request, StatusCode, Uri},
    response::{Html, IntoResponse, Response},
};
use std::path::PathBuf;
use tokio::fs;

use super::state::AppState;

/// 处理 ASP 请求
pub async fn handle_asp(uri: Uri, state: AppState) -> impl IntoResponse {
    let path = uri.path().trim_start_matches('/');
    let file_path = state.config.home_dir.join(path);

    // 检查文件是否存在
    if !file_path.exists() {
        return Html("File not found".to_string()).into_response();
    }

    // 读取文件内容
    match fs::read_to_string(&file_path).await {
        Ok(content) => {
            // 执行 ASP 文件
            let mut engine = crate::asp::Engine::new();
            match engine.execute(&content) {
                Ok(output) => Html(output).into_response(),
                Err(e) => Html(format!("<pre>Error: {}</pre>", e)).into_response(),
            }
        }
        Err(e) => Html(format!("<pre>Error reading file: {}</pre>", e)).into_response(),
    }
}

/// 处理静态文件请求
pub async fn handle_static(uri: Uri, state: AppState) -> impl IntoResponse {
    let path = uri.path().trim_start_matches('/');
    let file_path = state.config.home_dir.join(path);

    // 安全检查：防止路径遍历
    if !file_path.starts_with(&state.config.home_dir) {
        return Response::builder()
            .status(StatusCode::FORBIDDEN)
            .body(Body::from("Forbidden"))
            .unwrap();
    }

    // 检查是否是目录
    if file_path.is_dir() {
        // 尝试返回索引文件
        let index_path = file_path.join(&state.config.index_file);
        if index_path.exists() {
            return handle_asp(uri, state).await.into_response();
        }

        // 显示目录列表
        if state.config.directory_listing {
            return generate_directory_listing(&file_path, path).await;
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
            .body(Body::from("File not found"))
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
        Err(_) => Response::builder()
            .status(StatusCode::INTERNAL_SERVER_ERROR)
            .body(Body::from("Internal server error"))
            .unwrap(),
    }
}

/// 生成目录列表
async fn generate_directory_listing(dir: &PathBuf, url_path: &str) -> Response {
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
        let mut entries: Vec<_> = entries.filter_map(|e| e.ok()).collect();
        entries.sort_by_key(|e| e.file_name());

        for entry in entries {
            if let Ok(name) = entry.file_name().into_string() {
                let is_dir = entry.file_type().await.map(|t| t.is_dir()).unwrap_or(false);
                let display_name = if is_dir {
                    format!("{}/", name)
                } else {
                    name.clone()
                };
                html.push_str(&format!(
                    "<li><a href=\"{}\">{}</a></li>\n",
                    name, display_name
                ));
            }
        }
    }

    html.push_str("</ul>\n<hr>\n</body>\n</html>");
    Html(html).into_response()
}
