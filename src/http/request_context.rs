//! HTTP 请求上下文
//!
//! 存储 HTTP 请求的相关信息，传递给 ASP 引擎

use std::collections::HashMap;

/// HTTP 请求上下文
#[derive(Debug, Clone)]
pub struct RequestContext {
    /// 请求方法 (GET, POST, etc.)
    pub method: String,
    /// 请求路径
    pub path: String,
    /// 查询字符串参数
    pub query_string: HashMap<String, String>,
    /// 表单数据 (POST body)
    pub form_data: HashMap<String, String>,
    /// Cookies
    pub cookies: HashMap<String, String>,
    /// 服务器变量
    pub server_variables: HashMap<String, String>,
    /// 请求体 (原始)
    pub body: Vec<u8>,
}

impl Default for RequestContext {
    fn default() -> Self {
        Self::new()
    }
}

impl RequestContext {
    /// 创建新的请求上下文
    pub fn new() -> Self {
        RequestContext {
            method: "GET".to_string(),
            path: "/".to_string(),
            query_string: HashMap::new(),
            form_data: HashMap::new(),
            cookies: HashMap::new(),
            server_variables: HashMap::new(),
            body: Vec::new(),
        }
    }

    /// 解析查询字符串
    pub fn parse_query_string(query: &str) -> HashMap<String, String> {
        if query.is_empty() {
            return HashMap::new();
        }

        urlencoding::decode(query)
            .map(|decoded| {
                decoded
                    .split('&')
                    .filter_map(|pair| {
                        let parts: Vec<&str> = pair.splitn(2, '=').collect();
                        if parts.len() == 2 {
                            Some((parts[0].to_lowercase(), parts[1].to_string()))
                        } else if !parts[0].is_empty() {
                            Some((parts[0].to_lowercase(), String::new()))
                        } else {
                            None
                        }
                    })
                    .collect()
            })
            .unwrap_or_default()
    }

    /// 解析表单数据
    pub fn parse_form_data(body: &str) -> HashMap<String, String> {
        if body.is_empty() {
            return HashMap::new();
        }

        urlencoding::decode(body)
            .map(|decoded| {
                decoded
                    .split('&')
                    .filter_map(|pair| {
                        let parts: Vec<&str> = pair.splitn(2, '=').collect();
                        if parts.len() == 2 {
                            Some((parts[0].to_lowercase(), parts[1].to_string()))
                        } else if !parts[0].is_empty() {
                            Some((parts[0].to_lowercase(), String::new()))
                        } else {
                            None
                        }
                    })
                    .collect()
            })
            .unwrap_or_default()
    }

    /// 解析 Cookie 头
    pub fn parse_cookies(cookie_header: &str) -> HashMap<String, String> {
        if cookie_header.is_empty() {
            return HashMap::new();
        }

        cookie_header
            .split(';')
            .filter_map(|cookie| {
                let parts: Vec<&str> = cookie.trim().splitn(2, '=').collect();
                if parts.len() == 2 {
                    Some((parts[0].to_string(), parts[1].to_string()))
                } else {
                    None
                }
            })
            .collect()
    }

    /// 从 axum Request 构建请求上下文
    pub async fn from_request(request: axum::http::Request<axum::body::Body>) -> Self {
        use axum::http::{header, Method};

        let method = request.method().to_string();
        let uri = request.uri().clone();
        let path = uri.path().to_string();

        // 解析查询字符串
        let query_string = uri
            .query()
            .map(|q| Self::parse_query_string(q))
            .unwrap_or_default();

        // 解析 Cookies
        let cookies = request
            .headers()
            .get(header::COOKIE)
            .and_then(|v| v.to_str().ok())
            .map(|s| Self::parse_cookies(s))
            .unwrap_or_default();

        // 构建服务器变量
        let mut server_variables = HashMap::new();
        server_variables.insert("REQUEST_METHOD".to_string(), method.clone());
        server_variables.insert("PATH_INFO".to_string(), path.clone());
        server_variables.insert("QUERY_STRING".to_string(), uri.query().unwrap_or("").to_string());
        server_variables.insert("SERVER_PROTOCOL".to_string(), "HTTP/1.1".to_string());
        server_variables.insert("SERVER_NAME".to_string(), "localhost".to_string());
        server_variables.insert("SERVER_PORT".to_string(), "8080".to_string());

        // 添加 HTTP 头到服务器变量
        for (name, value) in request.headers() {
            let var_name = format!("HTTP_{}", name.as_str().to_uppercase().replace('-', "_"));
            if let Ok(v) = value.to_str() {
                server_variables.insert(var_name, v.to_string());
            }
        }

        // 读取请求体
        let (parts, body) = request.into_parts();
        let body_bytes = axum::body::to_bytes(body, 1024 * 1024) // 限制 1MB
            .await
            .unwrap_or_default()
            .to_vec();

        // 解析表单数据 (仅 POST 请求)
        let form_data = if method == Method::POST.as_str() {
            let content_type = parts
                .headers
                .get(header::CONTENT_TYPE)
                .and_then(|v| v.to_str().ok())
                .unwrap_or("");

            if content_type.starts_with("application/x-www-form-urlencoded") {
                String::from_utf8(body_bytes.clone())
                    .ok()
                    .map(|s| Self::parse_form_data(&s))
                    .unwrap_or_default()
            } else {
                HashMap::new()
            }
        } else {
            HashMap::new()
        };

        RequestContext {
            method,
            path,
            query_string,
            form_data,
            cookies,
            server_variables,
            body: body_bytes,
        }
    }

    /// 获取查询字符串参数
    pub fn query(&self, key: &str) -> Option<&String> {
        self.query_string.get(&key.to_lowercase())
    }

    /// 获取表单数据
    pub fn form(&self, key: &str) -> Option<&String> {
        self.form_data.get(&key.to_lowercase())
    }

    /// 获取 Cookie
    pub fn cookie(&self, key: &str) -> Option<&String> {
        self.cookies.get(&key.to_lowercase())
    }

    /// 获取服务器变量
    pub fn server_variable(&self, key: &str) -> Option<&String> {
        self.server_variables.get(&key.to_lowercase())
    }
}
