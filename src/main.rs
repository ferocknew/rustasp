use actix_files::NamedFile;
use actix_web::{web, App, HttpRequest, HttpResponse, HttpServer, Result};
use std::path::PathBuf;
use std::fs;

/// 配置结构体
struct Config {
    /// 是否显示目录列表
    directory_listing: bool,
    /// Web 根目录
    home_dir: PathBuf,
    /// 默认索引文件
    index_file: String,
    /// 服务端口
    port: u16,
}

impl Config {
    fn from_env() -> Self {
        // 加载 .env 文件
        dotenv::dotenv().ok();

        Config {
            directory_listing: std::env::var("DIRECTORY_LISTING")
                .unwrap_or_else(|_| "false".to_string())
                .parse()
                .unwrap_or(false),
            home_dir: PathBuf::from(
                std::env::var("HOME_DIR").unwrap_or_else(|_| "./www".to_string())
            ),
            index_file: std::env::var("INDEX_FILE")
                .unwrap_or_else(|_| "index.asp".to_string()),
            port: std::env::var("PORT")
                .unwrap_or_else(|_| "8080".to_string())
                .parse()
                .unwrap_or(8080),
        }
    }
}

/// 生成目录列表 HTML
fn generate_directory_listing(path: &PathBuf, url_path: &str) -> String {
    let mut html = format!(r#"<!DOCTYPE HTML>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Directory listing for {}</title>
</head>
<body>
<h1>Directory listing for {}</h1>
<hr>
<ul>
"#, url_path, url_path);

    if let Ok(entries) = fs::read_dir(path) {
        let mut entries: Vec<_> = entries.filter_map(|e| e.ok()).collect();
        entries.sort_by_key(|e| e.file_name());

        for entry in entries {
            let name = entry.file_name().to_string_lossy().to_string();
            let is_dir = entry.file_type().map(|t| t.is_dir()).unwrap_or(false);
            let display_name = if is_dir {
                format!("{}/", name)
            } else {
                name.clone()
            };
            html.push_str(&format!(r#"<li><a href="{}">{}</a></li>"#", name, display_name));
        }
    }

    html.push_str(r#"</ul>
<hr>
</body>
</html>"#);

    html
}

/// 处理 ASP 文件请求
async fn handle_asp(req: HttpRequest, config: web::Data<Config>) -> Result<HttpResponse> {
    let path: PathBuf = req.match_info().query("filename").parse().unwrap();
    let file_path = config.home_dir.join(&path);

    // 暂时返回文件内容，后续实现 VBScript 解释器
    if file_path.exists() {
        let content = std::fs::read_to_string(&file_path)
            .unwrap_or_else(|_| "Error reading file".to_string());
        Ok(HttpResponse::Ok()
            .content_type("text/html; charset=utf-8")
            .body(content))
    } else {
        Ok(HttpResponse::NotFound().body("File not found"))
    }
}

/// 处理目录和静态文件请求
async fn handle_path(req: HttpRequest, config: web::Data<Config>) -> Result<HttpResponse> {
    let path: PathBuf = req.match_info().query("path").parse().unwrap_or_default();
    let full_path = config.home_dir.join(&path);

    // 如果是目录
    if full_path.is_dir() {
        // 尝试返回默认索引文件
        let index_path = full_path.join(&config.index_file);
        if index_path.exists() {
            if config.index_file.ends_with(".asp") {
                // ASP 文件
                let content = std::fs::read_to_string(&index_path)
                    .unwrap_or_else(|_| "Error reading file".to_string());
                return Ok(HttpResponse::Ok()
                    .content_type("text/html; charset=utf-8")
                    .body(content));
            } else {
                // 静态文件
                if let Ok(file) = NamedFile::open(&index_path) {
                    return Ok(file.into_response(&req));
                }
            }
        }

        // 显示目录列表
        if config.directory_listing {
            let url_path = if path.as_os_str().is_empty() {
                "/".to_string()
            } else {
                format!("/{}", path.display())
            };
            let html = generate_directory_listing(&full_path, &url_path);
            return Ok(HttpResponse::Ok()
                .content_type("text/html; charset=utf-8")
                .body(html));
        }

        return Ok(HttpResponse::Forbidden().body("Directory listing is disabled"));
    }

    // 静态文件
    if full_path.exists() {
        let file = NamedFile::open(&full_path)?;
        Ok(file.into_response(&req))
    } else {
        Ok(HttpResponse::NotFound().body("File not found"))
    }
}

/// 根路径处理
async fn handle_root(config: web::Data<Config>) -> Result<HttpResponse> {
    let index_path = config.home_dir.join(&config.index_file);

    if index_path.exists() {
        let content = std::fs::read_to_string(&index_path)
            .unwrap_or_else(|_| "Error reading file".to_string());
        Ok(HttpResponse::Ok()
            .content_type("text/html; charset=utf-8")
            .body(content))
    } else {
        // 如果索引文件不存在且开启了目录列表
        if config.directory_listing {
            let html = generate_directory_listing(&config.home_dir, "/");
            return Ok(HttpResponse::Ok()
                .content_type("text/html; charset=utf-8")
                .body(html));
        }
        Ok(HttpResponse::NotFound().body("Index file not found"))
    }
}

#[actix_web::main]
async fn main() -> std::io::Result<()> {
    let config = Config::from_env();

    println!("🚀 VBScript ASP Server starting at http://127.0.0.1:{}", config.port);
    println!("📁 Home directory: {}", config.home_dir.display());
    println!("📄 Index file: {}", config.index_file);
    println!("📋 Directory listing: {}", config.directory_listing);

    let config_data = web::Data::new(config);

    HttpServer::new(move || {
        App::new()
            .app_data(config_data.clone())
            // 根路径
            .route("/", web::get().to(handle_root))
            // 处理 .asp 文件
            .route("/{filename:.*\\.asp}", web::get().to(handle_asp))
            // 处理其他路径
            .route("/{path:.*}", web::get().to(handle_path))
    })
    .bind(format!("127.0.0.1:{}", config_data.port))?
    .run()
    .await
}
