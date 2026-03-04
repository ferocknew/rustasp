//! HTTP 处理器

use axum::{
    body::Body,
    http::{StatusCode, Uri},
    response::{Html, IntoResponse, Response},
};
use std::path::PathBuf;
use tokio::fs;

use super::path_resolver::PathResolver;
use super::state::AppState;

/// 处理 ASP 请求（使用简化引擎）
pub async fn handle_asp(uri: Uri, state: AppState) -> impl IntoResponse {
    // 使用路径解析器安全解析路径
    let resolver = PathResolver::new(
        state.config.home_dir.clone(),
        state.config.allow_parent_paths,
    );

    let file_path = match resolver.resolve(uri.path()) {
        Ok(path) => path,
        Err(e) => {
            return Html(format!("<h1>403 - Forbidden</h1><pre>{}</pre>", e))
                .into_response();
        }
    };

    // 检查文件是否存在
    if !file_path.exists() {
        return Html("<h1>404 - File not found</h1>".to_string()).into_response();
    }

    // 读取文件内容
    match fs::read_to_string(&file_path).await {
        Ok(content) => {
            // 使用简化的 ASP 引擎执行
            let mut engine = crate::asp::Engine::new().with_debug(state.config.debug);
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
    // 使用路径解析器安全解析路径
    let resolver = PathResolver::new(
        state.config.home_dir.clone(),
        state.config.allow_parent_paths,
    );

    let file_path = match resolver.resolve(uri.path()) {
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
