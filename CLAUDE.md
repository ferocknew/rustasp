# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

这是一个 Rust 实现的 Classic ASP（VBScript 子集）运行时，目标是在不依赖 IIS 的情况下运行 ASP 应用，支持容器化部署。

## 历史背景（重要）

### 为什么有 SimpleEngine？

早期尝试使用 chumsky 解析器组合子实现完整的语句解析器，导致：
- **编译时间爆炸**：58 分钟编译时间
- **编译器内存占用巨大**

因此采用**渐进式策略**：
1. 保留完整的 Lexer（手写实现，快速编译）
2. 保留完整的 AST 定义
3. 保留完整的 Runtime
4. **SimpleEngine 作为过渡方案**
5. **渐进式接入语句解析**：一次只实现少量语句类型

### 渐进式接入原则

当接入语句解析器时：
- 每次只增加 1-2 种语句的解析
- 保证每次都能快速编译（< 1 分钟）
- 先从最常用的语句开始：Dim → If → For → While → Function

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

# 调试模式运行（显示 ASP 解析过程）
DEBUG=true cargo run

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
ASP 引擎层 (simple_engine / 完整引擎)
    ↓
VBScript Runtime (interpreter + value)
    ↓
Parser / AST (手写 Lexer + Pratt ExprParser)
```

## 目录结构

```
src/
├── main.rs                 # 入口
├── lib.rs                  # 库入口
│
├── ast/                    # 抽象语法树（纯数据结构）
│   ├── mod.rs
│   ├── op.rs               # 运算符定义
│   ├── expr.rs             # 表达式
│   ├── stmt.rs             # 语句（完整定义）
│   └── program.rs          # 程序
│
├── parser/                 # 语法解析层
│   ├── mod.rs
│   ├── keyword.rs          # 关键字定义
│   ├── lexer.rs            # 词法分析器（手写）✅
│   ├── expr_parser.rs      # 表达式解析器（Pratt）✅
│   └── error.rs            # 解析错误
│   ⚠️ 缺失：stmt_parser.rs（语句解析器）
│
├── runtime/                # 解释执行层 ✅
│   ├── mod.rs
│   ├── interpreter.rs      # 解释器调度
│   ├── context.rs          # 执行上下文
│   ├── scope.rs            # 变量作用域
│   ├── error.rs            # 运行时错误
│   └── value/              # 值类型系统
│
├── builtins/               # ASP 内建对象
│   ├── mod.rs
│   ├── response.rs         # Response 对象
│   ├── request.rs          # Request 对象
│   ├── server.rs           # Server 对象
│   └── session.rs          # Session 对象
│
├── asp/                    # ASP 引擎层
│   ├── mod.rs
│   └── simple_engine.rs    # 简化引擎（当前使用）
│
└── http/                   # HTTP 服务层
    ├── mod.rs
    ├── router.rs           # 路由配置
    ├── handler.rs          # 请求处理
    ├── state.rs            # 应用状态
    ├── path_resolver.rs    # 路径解析（安全防护）
    └── request_context.rs  # HTTP 请求上下文
```

## 模块职责

| 目录 | 职责 | 状态 |
|------|------|------|
| `src/ast/` | 抽象语法树定义，纯数据结构 | ✅ 完成 |
| `src/parser/lexer.rs` | 词法分析，Token 流生成 | ✅ 完成 |
| `src/parser/expr_parser.rs` | 表达式解析（Pratt 算法）| ✅ 完成 |
| `src/parser/stmt_parser.rs` | 语句解析 | ❌ 待实现 |
| `src/runtime/` | 解释执行 AST，变量作用域 | ✅ 完成 |
| `src/runtime/value/` | Value 类型系统 | ✅ 完成 |
| `src/builtins/` | ASP 内建对象 | ⚠️ 部分完成 |
| `src/asp/simple_engine.rs` | 简化 ASP 引擎 | ⚠️ 功能受限 |
| `src/http/` | Web 服务 | ✅ 完成 |

## 关键数据流

### 当前流程（SimpleEngine）

```
HTTP 请求
    ↓
http/handler.rs（加载 ASP 文件）
    ↓
http/request_context.rs（解析 QueryString/Form）
    ↓
asp/simple_engine.rs（简化解析）
    ↓
HTML 输出 → HTTP 响应
```

### 完整流程（待实现）

```
HTTP 请求
    ↓
http/handler.rs（加载 ASP 文件）
    ↓
http/request_context.rs（解析 QueryString/Form）
    ↓
parser/lexer.rs（源代码 → Token 流）
    ↓
parser/stmt_parser.rs（Token → AST）← 待实现
    ↓
runtime/interpreter.rs（执行 AST）
    ↓
builtins/（Response.Write 等）
    ↓
HTML 输出 → HTTP 响应
```

## 设计原则

### 单一职责
每个目录代表一个系统维度：语法、执行、内建对象、模板引擎、HTTP

### 控制文件大小
- **单文件不超过 500 行**
- match 分支过多必须拆文件
- Value 相关逻辑集中在 `runtime/value/` 目录

### 子集策略
支持：Dim, If, For, While, Function, 基本表达式, Response.Write

不支持：COM, ActiveX, Windows-only DLL

### 渐进式开发
- 一次只实现少量功能
- 确保编译时间可控
- 先核心功能，后扩展功能

## 配置

通过 `.env` 文件配置：

```env
# 是否显示目录列表
DIRECTORY_LISTING=true

# Web 根目录路径
HOME_DIR=./www

# 默认索引文件
INDEX_FILE=index.asp

# 服务端口号
PORT=8080

# 是否启用调试模式（显示 ASP 解析过程）
DEBUG=false

# 是否显示详细错误信息
DETAILED_ERROR=false

# 自定义错误页面路径（相对于 HOME_DIR）
# ERROR_PAGE=error.html

# ASP 文件扩展名（逗号分隔）
ASP_EXT=asp,asa

# 是否允许父路径访问（安全风险）
ALLOW_PARENT_PATH=false
```

## 技术栈

- Rust 2021 Edition
- axum 0.7 (Web 框架)
- tower / tower-http (中间件)
- tokio (异步运行时)
- serde (序列化)
- thiserror / anyhow (错误处理)
- urlencoding (URL 编解码)
- mime_guess (MIME 类型检测)

## 注意事项

### 关于 SimpleEngine

SimpleEngine 是一个过渡方案，它：
- 使用简单的字符串匹配和正则表达式
- 仅支持：`Response.Write`、`Request()`、`Dim`、简单变量赋值
- 不支持完整的 VBScript 语法（If、For、Function 等）

### 接入完整引擎

当需要接入完整引擎时：
1. 先实现 `stmt_parser.rs`，一次只解析 1-2 种语句
2. 在 `asp/mod.rs` 中切换引擎
3. 确保 Request/Response 对象与 Runtime 集成

### 编译时间控制

如果发现编译时间过长：
- 检查是否有复杂的泛型嵌套
- 检查是否有过多的 derive 宏
- 考虑拆分大文件
- 使用 `cargo check` 代替 `cargo build` 进行快速验证
