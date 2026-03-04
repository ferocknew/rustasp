//! 路由配置

use axum::{routing::get, Router};
use axum::response::{IntoResponse, Response};

use super::handler;
use super::state::AppState;

/// 创建路由
pub fn create_router(state: AppState) -> Router {
    Router::new()
        .route("/", get(root_handler))
        .route("/*path", get(path_handler))
        .with_state(state)
}

/// 根路径处理器
async fn root_handler(
    axum::extract::State(state): axum::extract::State<AppState>,
) -> Response {
    // 尝试访问索引文件
    let index_path = state.config.home_dir.join(&state.config.index_file);

    if index_path.exists() {
        // 索引文件存在，执行它
        let uri = axum::http::Uri::builder()
            .path_and_query(&format!("/{}", state.config.index_file))
            .build()
            .unwrap();
        handler::handle_asp(uri, state).await.into_response()
    } else {
        // 索引文件不存在，显示目录列表或返回 403
        if state.config.directory_listing {
            handler::generate_directory_listing(&state.config.home_dir, "/").await
        } else {
            Response::builder()
                .status(403)
                .body(axum::body::Body::from("Directory listing is disabled"))
                .unwrap()
        }
    }
}

/// 路径处理器
async fn path_handler(
    axum::extract::State(state): axum::extract::State<AppState>,
    axum::extract::Path(path): axum::extract::Path<String>,
) -> Response {
    let uri = axum::http::Uri::builder()
        .path_and_query(&format!("/{}", path))
        .build()
        .unwrap();

    // ASP 文件处理
    if path.ends_with(".asp") || path.ends_with(".asa") {
        return handler::handle_asp(uri, state).await.into_response();
    }

    // 静态文件处理
    handler::handle_static(uri, state).await.into_response()
}
