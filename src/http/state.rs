//! 应用状态

use std::path::PathBuf;

/// 应用配置
#[derive(Debug, Clone)]
pub struct Config {
    /// Web 根目录
    pub home_dir: PathBuf,
    /// 默认索引文件
    pub index_file: String,
    /// 服务端口
    pub port: u16,
    /// 是否显示目录列表
    pub directory_listing: bool,
    /// 是否启用调试模式
    pub debug: bool,
    /// 是否允许父路径访问 (../)
    pub allow_parent_paths: bool,
}

impl Default for Config {
    fn default() -> Self {
        Config {
            home_dir: PathBuf::from("./www"),
            index_file: "index.asp".to_string(),
            port: 8080,
            directory_listing: false,
            debug: false,
            allow_parent_paths: false,
        }
    }
}

impl Config {
    /// 从环境变量加载配置
    pub fn from_env() -> Self {
        dotenv::dotenv().ok();

        Config {
            home_dir: PathBuf::from(
                std::env::var("HOME_DIR").unwrap_or_else(|_| "./www".to_string()),
            ),
            index_file: std::env::var("INDEX_FILE").unwrap_or_else(|_| "index.asp".to_string()),
            port: std::env::var("PORT")
                .unwrap_or_else(|_| "8080".to_string())
                .parse()
                .unwrap_or(8080),
            directory_listing: std::env::var("DIRECTORY_LISTING")
                .unwrap_or_else(|_| "false".to_string())
                .parse()
                .unwrap_or(false),
            debug: std::env::var("DEBUG")
                .unwrap_or_else(|_| "false".to_string())
                .parse()
                .unwrap_or(false),
            allow_parent_paths: std::env::var("ALLOW_PARENT_PATH")
                .unwrap_or_else(|_| "false".to_string())
                .parse()
                .unwrap_or(false),
        }
    }
}

/// 应用状态
#[derive(Debug, Clone)]
pub struct AppState {
    pub config: Config,
}
