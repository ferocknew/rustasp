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
    /// 是否显示详细错误信息
    pub detailed_error: bool,
    /// 自定义错误页面（相对于 home_dir）
    pub error_page: Option<String>,
    /// ASP 文件扩展名（逗号分隔）
    pub asp_ext: Vec<String>,
    /// 索引文件功能开关
    pub index_file_enable: bool,
    /// Runtime 目录路径（用于 Session、缓存等硬盘存储）
    pub runtime_dir: PathBuf,
    /// Now 函数日期时间格式
    #[allow(dead_code)]
    pub now_format: String,
    /// Date 函数日期格式
    #[allow(dead_code)]
    pub date_format: String,
    /// Time 函数时间格式
    #[allow(dead_code)]
    pub time_format: String,
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
            detailed_error: false,
            error_page: None,
            asp_ext: vec!["asp".to_string(), "asa".to_string()],
            index_file_enable: true,
            runtime_dir: PathBuf::from("./runtime"),
            now_format: "YYYY/MM/DD HH:MM:SS".to_string(),
            date_format: "YYYY/MM/DD".to_string(),
            time_format: "HH:MM:SS".to_string(),
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
            detailed_error: std::env::var("DETAILED_ERROR")
                .unwrap_or_else(|_| "false".to_string())
                .parse()
                .unwrap_or(false),
            error_page: std::env::var("ERROR_PAGE").ok(),
            asp_ext: std::env::var("ASP_EXT")
                .unwrap_or_else(|_| "asp,asa".to_string())
                .split(',')
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty())
                .collect(),
            index_file_enable: std::env::var("INDEX_FILE_ENABLE")
                .unwrap_or_else(|_| "true".to_string())
                .parse()
                .unwrap_or(true),
            runtime_dir: PathBuf::from(
                std::env::var("RUNTIME_DIR").unwrap_or_else(|_| "./runtime".to_string()),
            ),
            now_format: std::env::var("NOW_FORMAT")
                .unwrap_or_else(|_| "YYYY/MM/DD HH:MM:SS".to_string()),
            date_format: std::env::var("DATE_FORMAT")
                .unwrap_or_else(|_| "YYYY/MM/DD".to_string()),
            time_format: std::env::var("TIME_FORMAT")
                .unwrap_or_else(|_| "HH:MM:SS".to_string()),
        }
    }

    /// 获取 Session 存储类型
    #[allow(dead_code)]
    pub fn session_storage_type(&self) -> String {
        std::env::var("SESSION_STORAGE").unwrap_or_else(|_| "memory".to_string())
    }

    /// 获取 Session 超时时间（分钟）
    #[allow(dead_code)]
    pub fn session_timeout(&self) -> u32 {
        std::env::var("SESSION_TIMEOUT")
            .unwrap_or_else(|_| "20".to_string())
            .parse()
            .unwrap_or(20)
    }
}

/// 应用状态
#[derive(Debug, Clone)]
pub struct AppState {
    pub config: Config,
}
