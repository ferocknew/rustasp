# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是一个 Rust 实现的 Classic ASP（VBScript 子集）运行时，目标是在不依赖 IIS 的情况下运行 ASP 应用，支持容器化部署。

## 常用命令

```bash
# 快速类型检查
cargo check

# 开发运行
cargo run

# Release 编译
cargo run --release

# 带日志运行
RUST_LOG=debug cargo run

# 运行测试
cargo test

# 运行特定测试
cargo test test_name

# 显示测试输出
cargo test -- --nocapture

# 代码检查
cargo clippy

# 格式化
cargo fmt
```

## 架构分层

```
HTTP 层 (axum)
    ↓
ASP 引擎层 (segmenter + engine)
    ↓
VBScript Runtime (interpreter + value)
    ↓
Parser / AST (chumsky)
```

## 模块职责

| 目录 | 职责 |
|------|------|
| `src/ast/` | 抽象语法树定义，纯数据结构，不包含执行逻辑 |
| `src/parser/` | 词法分析 (lexer) + 语法分析 (parser)，使用 chumsky 解析器组合子 |
| `src/runtime/` | 解释执行 AST，管理变量作用域和函数调用，实现弱类型系统 |
| `src/runtime/value/` | Value 类型系统，包含比较、转换、运算、显示格式化 |
| `src/builtins/` | ASP 内建对象：Response、Request、Server、Session |
| `src/asp/` | ASP 引擎：分割 HTML 与 `<% %>`，执行脚本片段，拼接输出 |
| `src/http/` | Web 服务：路由配置、请求处理、应用状态，使用 axum |

## 关键数据流

1. HTTP 请求 → `http/handler.rs` 加载 ASP 文件
2. `asp/segmenter.rs` 将文件分割为 HTML 和代码段
3. `parser/` 将代码段解析为 AST
4. `runtime/interpreter.rs` 执行 AST
5. `builtins/` 提供 Response.Write 等内建功能
6. 输出返回 HTTP 响应

## 设计原则

- **单文件不超过 500 行**：match 分支过多时拆文件
- **Value 相关逻辑集中在 `runtime/value/` 目录**
- **子集策略**：支持 Dim, If, Function, 基本表达式；不支持 COM, ActiveX, Windows-only DLL

## 配置

通过 `.env` 文件配置：
- `DIRECTORY_LISTING` - 是否显示目录列表
- `HOME_DIR` - Web 根目录路径
- `INDEX_FILE` - 默认索引文件
- `PORT` - 服务端口号

## 技术栈

- Rust 2021 Edition
- axum 0.7 (Web 框架)
- tower / tower-http (中间件)
- chumsky 0.9 (解析器组合子)
- tokio (异步运行时)
- serde (序列化)
- thiserror / anyhow (错误处理)
