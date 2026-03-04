use actix_files::NamedFile;
use actix_web::{web, App, HttpRequest, HttpResponse, HttpServer, Result};
use std::path::PathBuf;

/// 处理 ASP 文件请求
async fn handle_asp(req: HttpRequest) -> Result<HttpResponse> {
    let path: PathBuf = req.match_info().query("filename").parse().unwrap();
    let file_path = PathBuf::from("www").join(&path);

    // 暂时返回文件内容，后续实现 VBScript 解释器
    if file_path.exists() {
        let content = std::fs::read_to_string(&file_path)
            .unwrap_or_else(|_| "Error reading file".to_string());
        Ok(HttpResponse::Ok()
            .content_type("text/html; charset=utf-8")
            .body(format!("<!-- ASP File: {} -->\n{}", path.display(), content)))
    } else {
        Ok(HttpResponse::NotFound().body("File not found"))
    }
}

/// 处理静态文件请求
async fn handle_static(req: HttpRequest) -> Result<NamedFile> {
    let path: PathBuf = req.match_info().query("path").parse().unwrap();
    let file_path = PathBuf::from("www").join(&path);
    Ok(NamedFile::open(file_path)?)
}

#[actix_web::main]
async fn main() -> std::io::Result<()> {
    println!("🚀 VBScript ASP Server starting at http://127.0.0.1:8080");

    HttpServer::new(|| {
        App::new()
            // 处理 .asp 文件
            .route("/{filename:.*\\.asp}", web::get().to(handle_asp))
            // 静态文件服务
            .route("/{path:.*}", web::get().to(handle_static))
            // 根路径
            .route("/", web::get().to(|| async {
                HttpResponse::Found()
                    .insert_header(("Location", "/index.asp"))
                    .finish()
            }))
    })
    .bind("127.0.0.1:8080")?
    .run()
    .await
}
