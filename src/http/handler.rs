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
use vbscript::runtime::objects::SessionManager;

/// 处理 ASP 请求
pub async fn handle_asp(uri: Uri, state: AppState, request: Request<Body>) -> impl IntoResponse {
    let uri = uri;
    let uri_str = uri.path();

    // 记录 HTTP 请求信息
    if state.config.debug {
        let method = request.method();
        let headers = request.headers();
        let user_agent = headers
            .get("user-agent")
            .and_then(|v| v.to_str().ok())
            .unwrap_or("Unknown");
        let content_type = headers
            .get("content-type")
            .and_then(|v| v.to_str().ok())
            .unwrap_or("None");

        println!("\n🌐 HTTP Request: {} {}", method, uri_str);
        println!("   User-Agent: {}", user_agent);
        println!("   Content-Type: {}", content_type);
    }

    let error_gen = ErrorPageGenerator::from_state(&state).await;
    let request_ctx = RequestContext::from_request(request).await;

    // 解析路径
    let resolver = PathResolver::new(
        state.config.home_dir.clone(),
        state.config.allow_parent_paths,
    );
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

    // 记录 HTTP 请求信息
    if state.config.debug {
        let method = request.method();
        println!("\n🌐 HTTP Request (Static): {} {}", method, uri_str);
    }

    let error_gen = ErrorPageGenerator::from_state(&state).await;

    // 解析路径
    let resolver = PathResolver::new(
        state.config.home_dir.clone(),
        state.config.allow_parent_paths,
    );
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
        // 尝试索引文件（仅在启用时）
        if state.config.index_file_enable {
            for index_name in state.config.index_file.split(',') {
                let index_name = index_name.trim();
                if index_name.is_empty() {
                    continue;
                }
                let index_path = file_path.join(index_name);
                if index_path.exists() {
                    return handle_asp(uri, state, request).await.into_response();
                }
            }
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
    // 尝试索引文件（仅在启用时）
    if state.config.index_file_enable {
        // 支持多个索引文件（逗号分隔）
        for index_name in state.config.index_file.split(',') {
            let index_name = index_name.trim();
            if index_name.is_empty() {
                continue;
            }
            let index_path = dir_path.join(index_name);
            if index_path.exists() {
                return execute_asp_file(&index_path, uri_str, state, request_ctx, error_gen).await;
            }
        }
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
    let processed_content =
        match crate::asp::preprocess(&content, file_path, &state.config.home_dir) {
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

    // 创建 SessionManager
    let session_manager = match SessionManager::new(&state.config.runtime_dir) {
        Ok(sm) => sm,
        Err(e) => {
            eprintln!("警告: 无法创建 SessionManager: {}", e);
            // 没有 SessionManager 也能继续，只是没有 Session 持久化
            return execute_asp_file_without_session(
                file_path,
                uri_str,
                state,
                request_ctx,
                error_gen,
            )
            .await;
        }
    };

    // 获取或创建 Session ID
    let _session_id = if let Some(existing_id) = request_ctx.cookie("ASPSESSIONID") {
        existing_id.to_string()
    } else {
        // 生成新的 Session ID
        SessionManager::generate_session_id()
    };

    let mut engine = crate::asp::Engine::new()
        .with_debug(state.config.debug)
        .with_request_context(request_ctx.clone())
        .with_session_manager(session_manager)
        .with_home_dir(state.config.home_dir.to_string_lossy().to_string());

    match engine.execute(&processed_content) {
        Ok(result) => {
            // 调试信息：显示执行结果
            if state.config.debug {
                println!("   ✅ ASP 执行成功");
                println!(
                    "   📊 输出长度: {} bytes, 状态码: {}",
                    result.output.len(),
                    result.response.get_status()
                );
            }

            // 检查是否是重定向
            if result.response.get_status() == 302 {
                if let Some(location) = result.response.get_headers().get("Location") {
                    // 在重定向响应中构建响应（Cookie 已在 engine.rs 中设置）
                    let mut builder = AxumResponse::builder()
                        .status(StatusCode::FOUND)
                        .header("Location", location.clone());

                    // 添加自定义响应头（包括 Cookie）
                    for (name, value) in result.response.get_headers() {
                        if name != "Location" {
                            builder = builder.header(name, value);
                        }
                    }

                    return builder.body(Body::empty()).unwrap();
                }
            }

            // 构建响应
            let mut builder = AxumResponse::builder();

            // 设置状态码
            let status =
                StatusCode::from_u16(result.response.get_status()).unwrap_or(StatusCode::OK);
            builder = builder.status(status);

            // 设置 Content-Type
            let content_type = result.response.get_content_type();
            builder = builder.header("Content-Type", content_type);

            // 添加自定义响应头（包括 Session Cookie）
            for (name, value) in result.response.get_headers() {
                if name != "Location" {
                    // Location 已在重定向中处理
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
            // 调试信息：显示执行错误
            if state.config.debug {
                println!("   ❌ ASP 执行失败: {}", e);
            }
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

/// 执行 ASP 文件（无 Session 支持，用于 SessionManager 初始化失败时）
async fn execute_asp_file_without_session(
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
    let processed_content =
        match crate::asp::preprocess(&content, file_path, &state.config.home_dir) {
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
        .with_request_context(request_ctx.clone())
        .with_home_dir(state.config.home_dir.to_string_lossy().to_string());

    match engine.execute(&processed_content) {
        Ok(result) => {
            // 调试信息：显示执行结果
            if state.config.debug {
                println!("   ✅ ASP 执行成功 (无 Session)");
                println!(
                    "   📊 输出长度: {} bytes, 状态码: {}",
                    result.output.len(),
                    result.response.get_status()
                );
            }

            // 构建响应
            let mut builder = AxumResponse::builder();

            // 设置状态码
            let status =
                StatusCode::from_u16(result.response.get_status()).unwrap_or(StatusCode::OK);
            builder = builder.status(status);

            // 设置 Content-Type
            let content_type = result.response.get_content_type();
            builder = builder.header("Content-Type", content_type);

            // 添加自定义响应头
            for (name, value) in result.response.get_headers() {
                builder = builder.header(name, value);
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
            let display = if is_dir {
                format!("{}/", name)
            } else {
                name.clone()
            };
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
