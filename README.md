# VBScript ASP Server

一个用 Rust 语言实现的 VBScript 解释器，提供 Classic ASP 支持。

## 当前功能

### HTTP 服务器
- ASP 文件路由：`.asp` 文件请求会被处理
- 静态文件服务：其他文件直接返回
- 目录列表：可配置是否显示目录内容
- 配置支持：通过 `.env` 文件配置

## 配置

在项目根目录创建 `.env` 文件：

```env
# 是否显示目录列表 (true/false)
DIRECTORY_LISTING=true

# Web 根目录路径
HOME_DIR=./www

# 默认索引文件
INDEX_FILE=index.asp

# 服务端口号
PORT=8080
```

## 快速开始

```bash
# 编译项目
cargo build

# 运行服务器（开发模式）
cargo run

# 运行服务器（ release 模式，性能更好）
cargo run --release
```

服务器启动后会显示配置信息：
```
🚀 VBScript ASP Server starting at http://127.0.0.1:8080
📁 Home directory: ./www
📄 Index file: index.asp
📋 Directory listing: true
```

## 使用示例

```bash
# 访问 ASP 文件
curl http://127.0.0.1:8080/index.asp

# 访问根目录（返回索引文件或目录列表）
curl http://127.0.0.1:8080/

# 浏览器访问
open http://127.0.0.1:8080/
```

## 项目结构

```
rust_vbscript/
├── src/
│   └── main.rs      # HTTP 服务器入口
├── www/             # Web 根目录
│   └── index.asp    # 示例 ASP 文件
├── .env             # 配置文件
└── Cargo.toml
```

## 调试

```bash
# 编译并显示详细输出
cargo build --verbose

# 运行并查看日志
RUST_LOG=debug cargo run

# 检查代码问题
cargo check

# 运行测试（如果有）
cargo test
```

## 技术栈

- Rust 2021 Edition
- actix-web 4.x - Web 框架
- actix-files - 静态文件服务
- tokio - 异步运行时
- dotenv - 环境变量配置

## 开发计划

- [x] HTTP 服务器基础框架
- [x] 配置文件支持 (.env)
- [x] 目录列表功能
- [ ] VBScript 词法分析器
- [ ] VBScript 语法解析器
- [ ] VBScript 解释器
- [ ] ASP 内置对象 (Response, Request, Session, Application 等)
- [ ] 数据库连接支持 (ADO)
