//! VBScript ASP Server
//!
//! 一个用 Rust 实现的 Classic ASP 子集运行时

mod asp;
mod ast;
mod builtins;
mod http;
mod parser;
mod runtime;

use http::create_router;
use http::state::Config;

#[tokio::main]
async fn main() {
    // 加载配置
    let config = Config::from_env();

    // 打印启动信息
    println!(
        "🚀 VBScript ASP Server starting at http://127.0.0.1:{}",
        config.port
    );
    println!("📁 Home directory: {}", config.home_dir.display());
    println!("📄 Index file: {}", config.index_file);
    println!("📋 Directory listing: {}", config.directory_listing);

    // 创建应用状态
    let state = http::AppState { config };

    // 创建路由
    let app = create_router(state);

    // 启动服务器
    let listener = tokio::net::TcpListener::bind("127.0.0.1:8080")
        .await
        .unwrap();
    axum::serve(listener, app).await.unwrap();
}
