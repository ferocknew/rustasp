# VBScript ASP Server

一个用 Rust 语言实现的 VBScript 解释器，提供 Classic ASP 支持。

## 当前功能

### HTTP 服务器
- 监听地址：`127.0.0.1:8080`
- ASP 文件路由：`.asp` 文件请求会被处理
- 静态文件服务：其他文件直接返回
- 根路径重定向：`/` 自动跳转到 `/index.asp`

## 快速开始

```bash
# 运行服务器
cargo run

# 访问 ASP 文件
curl http://127.0.0.1:8080/index.asp
```

## 项目结构

```
rust_vbscript/
├── src/
│   └── main.rs      # HTTP 服务器入口
├── www/             # Web 根目录
│   └── index.asp    # 示例 ASP 文件
└── Cargo.toml
```

## 技术栈

- Rust 2021 Edition
- actix-web 4.x - Web 框架
- actix-files - 静态文件服务
- tokio - 异步运行时

## 开发计划

- [x] HTTP 服务器基础框架
- [ ] VBScript 词法分析器
- [ ] VBScript 语法解析器
- [ ] VBScript 解释器
- [ ] ASP 内置对象 (Response, Request, Session, Application 等)
- [ ] 数据库连接支持 (ADO)
