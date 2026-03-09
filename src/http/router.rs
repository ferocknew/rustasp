//! 路由配置

use axum::routing::get;
use axum::response::{IntoResponse, Response};

use super::handler;
use super::state::AppState;

/// 创建路由
pub fn create_router(state: AppState) -> Router {
    Router::new()
        .route("/", get(root_handler).post(root_handler))
        .route("/*path", get(path_handler).post(path_handler))
        .with_state(state)
}

use axum::Router;
use axum::extract::Request;
use axum::body::Body;

/// 根路径处理器
async fn root_handler(
    axum::extract::State(state): axum::extract::State<AppState>,
    request: Request<Body>,
) -> Response {
    // 检查是否启用索引文件功能
    if !state.config.index_file_enable {
        // 索引文件功能已禁用，显示目录列表或返回 403
        if state.config.directory_listing {
            return handler::generate_directory_listing(&state.config.home_dir, "/").await;
        } else {
            return Response::builder()
                .status(403)
                .body(axum::body::Body::from("Directory listing is disabled"))
                .unwrap();
        }
    }

    // 尝试访问索引文件（支持多个索引文件，逗号分隔）
    for index_name in state.config.index_file.split(',') {
        let index_name = index_name.trim();
        if index_name.is_empty() {
            continue;
        }

        let index_path = state.config.home_dir.join(index_name);

        if index_path.exists() {
            // 检查索引文件是否是 ASP 文件
            let is_asp = state.config.asp_ext.iter().any(|ext| {
                index_name.to_lowercase().ends_with(&format!(".{}", ext.to_lowercase()))
            });

            let uri = axum::http::Uri::builder()
                .path_and_query(&format!("/{}", index_name))
                .build()
                .unwrap();

            if is_asp {
                // ASP 文件执行
                return handler::handle_asp(uri, state, request).await.into_response();
            } else {
                // 静态文件直接返回
                return handler::handle_static(uri, state, request).await.into_response();
            }
        }
    }

    // 所有索引文件都不存在，显示目录列表或返回 403
    if state.config.directory_listing {
        handler::generate_directory_listing(&state.config.home_dir, "/").await
    } else {
        Response::builder()
            .status(403)
            .body(axum::body::Body::from("Directory listing is disabled"))
            .unwrap()
    }
}

/// 路径处理器
async fn path_handler(
    axum::extract::State(state): axum::extract::State<AppState>,
    axum::extract::Path(path): axum::extract::Path<String>,
    request: Request<Body>,
) -> Response {
    let uri = axum::http::Uri::builder()
        .path_and_query(&format!("/{}", path))
        .build()
        .unwrap();

    // 检查是否是 ASP 文件（根据配置的扩展名）
    let is_asp = state.config.asp_ext.iter().any(|ext| {
        path.to_lowercase().ends_with(&format!(".{}", ext.to_lowercase()))
    });

    if is_asp {
        return handler::handle_asp(uri, state, request).await.into_response();
    }

    // 静态文件处理
    handler::handle_static(uri, state, request).await.into_response()
}
