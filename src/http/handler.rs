//! HTTP 处理器

use axum::{
    body::Body,
    http::{Request, StatusCode, Uri},
    response::{Html, IntoResponse, Response as AxumResponse},
};
use std::path::PathBuf;
use tokio::fs;

use super::error_page::{ErrorInfo, ErrorKind, ErrorPageGenerator};
use super::path_resolver::PathResolver;
use super::request_context::RequestContext;
use super::state::AppState;

/// 处理 ASP 请求
pub async fn handle_asp(uri: Uri, state: AppState, request: Request<Body>) -> impl IntoResponse {
    let uri_str = uri.path();
    let error_gen = ErrorPageGenerator::from_state(&state).await;
    let request_ctx = RequestContext::from_request(request).await;

    // 解析路径
    let resolver = PathResolver::new(state.config.home_dir.clone(), state.config.allow_parent_paths);
    let file_path = match resolver.resolve(uri_str) {
        Ok(path) => path,
        Err(e) => {
            return error_gen.generate(&ErrorInfo::new(
                "handler.rs",
                0,
                uri_str,
                "(path resolution failed)",
                ErrorKind::PathResolution,
                e.to_string(),
            ));
        }
    };

    // 检查文件存在性
    if !file_path.exists() {
        return error_gen.generate(&ErrorInfo::new(
            "handler.rs",
            0,
            uri_str,
            file_path.display().to_string(),
            ErrorKind::FileNotFound,
            "The requested file does not exist.",
        ));
    }

    // 处理目录
    if file_path.is_dir() {
        return handle_directory(&file_path, uri_str, &state, &request_ctx, &error_gen).await;
    }

    // 执行 ASP 文件
    execute_asp_file(&file_path, uri_str, &state, &request_ctx, &error_gen).await
}

/// 处理静态文件请求
pub async fn handle_static(uri: Uri, state: AppState, request: Request<Body>) -> impl IntoResponse {
    let uri_str = uri.path();
    let error_gen = ErrorPageGenerator::from_state(&state).await;

    // 解析路径
    let resolver = PathResolver::new(state.config.home_dir.clone(), state.config.allow_parent_paths);
    let file_path = match resolver.resolve(uri_str) {
        Ok(path) => path,
        Err(e) => {
            return AxumResponse::builder()
                .status(StatusCode::FORBIDDEN)
                .body(Body::from(format!("Forbidden: {}", e)))
                .unwrap();
        }
    };

    // 处理目录
    if file_path.is_dir() {
        let index_path = file_path.join(&state.config.index_file);
        if index_path.exists() {
            return handle_asp(uri, state, request).await.into_response();
        }
        if state.config.directory_listing {
            return generate_directory_listing(&file_path, uri.path()).await;
        }
        return AxumResponse::builder()
            .status(StatusCode::FORBIDDEN)
            .body(Body::from("Directory listing is disabled"))
            .unwrap();
    }

    // 检查文件存在性
    if !file_path.exists() {
        return error_gen.generate(&ErrorInfo::new(
            "handler.rs",
            0,
            uri_str,
            file_path.display().to_string(),
            ErrorKind::FileNotFound,
            "The requested file does not exist.",
        ));
    }

    // 读取并返回文件
    match fs::read(&file_path).await {
        Ok(content) => {
            let mime = mime_guess::from_path(&file_path)
                .first_or_octet_stream()
                .to_string();
            AxumResponse::builder()
                .status(StatusCode::OK)
                .header("Content-Type", mime)
                .body(Body::from(content))
                .unwrap()
        }
        Err(e) => error_gen.generate(&ErrorInfo::new(
            "handler.rs",
            0,
            uri_str,
            file_path.display().to_string(),
            ErrorKind::FileRead,
            e.to_string(),
        )),
    }
}

/// 处理目录请求
async fn handle_directory(
    dir_path: &PathBuf,
    uri_str: &str,
    state: &AppState,
    request_ctx: &RequestContext,
    error_gen: &ErrorPageGenerator,
) -> AxumResponse {
    // 尝试索引文件
    let index_path = dir_path.join(&state.config.index_file);
    if index_path.exists() {
        return execute_asp_file(&index_path, uri_str, state, request_ctx, error_gen).await;
    }

    // 目录列表
    if state.config.directory_listing {
        return generate_directory_listing(dir_path, uri_str).await;
    }

    error_gen.generate(&ErrorInfo::new(
        "handler.rs",
        0,
        uri_str,
        dir_path.display().to_string(),
        ErrorKind::DirectoryListingDisabled,
        "Directory listing is disabled and no index file found.",
    ))
}

/// 执行 ASP 文件
async fn execute_asp_file(
    file_path: &PathBuf,
    uri_str: &str,
    state: &AppState,
    request_ctx: &RequestContext,
    error_gen: &ErrorPageGenerator,
) -> AxumResponse {
    let content = match fs::read_to_string(file_path).await {
        Ok(c) => c,
        Err(e) => {
            return error_gen.generate(&ErrorInfo::new(
                "handler.rs",
                0,
                uri_str,
                file_path.display().to_string(),
                ErrorKind::FileRead,
                e.to_string(),
            ));
        }
    };

    // 预处理 include 指令
    let processed_content = match crate::asp::preprocess(&content, file_path, &state.config.home_dir) {
        Ok(c) => c,
        Err(e) => {
            return error_gen.generate(&ErrorInfo::new(
                "handler.rs",
                0,
                uri_str,
                file_path.display().to_string(),
                ErrorKind::AspExecution,
                format!("Include error: {}", e),
            ));
        }
    };

    let mut engine = crate::asp::Engine::new()
        .with_debug(state.config.debug)
        .with_request_context(request_ctx.clone());

    match engine.execute(&processed_content) {
        Ok(result) => {
            // 检查是否是重定向
            if result.response.get_status() == 302 {
                if let Some(location) = result.response.get_headers().get("Location") {
                    return AxumResponse::builder()
                        .status(StatusCode::FOUND)
                        .header("Location", location.clone())
                        .body(Body::empty())
                        .unwrap();
                }
            }

            // 构建响应
            let mut builder = AxumResponse::builder();

            // 设置状态码
            let status = StatusCode::from_u16(result.response.get_status()).unwrap_or(StatusCode::OK);
            builder = builder.status(status);

            // 设置 Content-Type
            let content_type = result.response.get_content_type();
            builder = builder.header("Content-Type", content_type);

            // 添加自定义响应头
            for (name, value) in result.response.get_headers() {
                if name != "Location" {  // Location 已在重定向中处理
                    builder = builder.header(name, value);
                }
            }

            // 返回响应体
            let output = result.output;
            match builder.body(Body::from(output.clone())) {
                Ok(resp) => resp,
                Err(_) => Html(output).into_response(),
            }
        }
        Err(e) => {
            let error_info = ErrorInfo::new(
                "engine.rs",
                0,
                uri_str,
                file_path.display().to_string(),
                ErrorKind::AspExecution,
                e.to_string(),
            )
            .with_source_code(&processed_content);
            error_gen.generate(&error_info)
        }
    }
}

/// 生成目录列表
pub async fn generate_directory_listing(dir: &PathBuf, url_path: &str) -> AxumResponse {
    let mut entries = match fs::read_dir(dir).await {
        Ok(entries) => entries,
        Err(_) => return Html("<h1>Cannot read directory</h1>".to_string()).into_response(),
    };

    let mut items = Vec::new();
    while let Ok(Some(entry)) = entries.next_entry().await {
        if let Ok(name) = entry.file_name().into_string() {
            let is_dir = entry.file_type().await.map(|t| t.is_dir()).unwrap_or(false);
            items.push((name, is_dir));
        }
    }
    items.sort_by(|a, b| a.0.cmp(&b.0));

    let base = url_path.trim_end_matches('/');
    let list_items: String = items
        .into_iter()
        .map(|(name, is_dir)| {
            let display = if is_dir { format!("{}/", name) } else { name.clone() };
            let href = if url_path.is_empty() || url_path == "/" {
                format!("/{}", name)
            } else {
                format!("{}/{}", base, name)
            };
            format!("<li><a href=\"{}\">{}</a></li>", href, display)
        })
        .collect();

    Html(format!(
        r#"<!DOCTYPE HTML>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Directory listing for {}</title>
</head>
<body>
<h1>Directory listing for {}</h1>
<hr>
<ul>
{}</ul>
<hr>
</body>
</html>"#,
        url_path, url_path, list_items
    ))
    .into_response()
}
