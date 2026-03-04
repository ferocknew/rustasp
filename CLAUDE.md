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

# 运行解析器测试工具
cargo run --bin test_parser
```

## 架构分层

```
HTTP 层 (axum)
    ↓
ASP 引擎层 (segmenter + engine + include)
    ↓
VBScript Runtime (interpreter + value)
    ↓
Parser / AST (手写 Lexer + Pratt ExprParser + StmtParser)
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
│   ├── stmt.rs             # 语句
│   └── program.rs          # 程序
│
├── parser/                 # 语法解析层
│   ├── mod.rs
│   ├── keyword.rs          # 关键字定义
│   ├── lexer.rs            # 词法分析器（手写）
│   ├── expr_parser.rs      # 表达式解析器（Pratt 算法）
│   ├── error.rs            # 解析错误
│   └── stmt_parser/        # 语句解析器（递归下降）
│       ├── mod.rs
│       ├── core.rs         # 核心解析逻辑
│       ├── control.rs      # 控制流语句
│       ├── declarations.rs # 声明语句
│       ├── procedures.rs   # 函数/过程
│       └── helpers.rs      # 辅助函数
│
├── runtime/                # 解释执行层
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
│   ├── engine.rs           # ASP 执行引擎
│   ├── segmenter.rs        # 代码分段器
│   └── include.rs          # Include 指令处理
│
└── http/                   # HTTP 服务层
    ├── mod.rs
    ├── router.rs           # 路由配置
    ├── handler.rs          # 请求处理
    ├── state.rs            # 应用状态
    ├── path_resolver.rs    # 路径解析（安全防护）
    ├── request_context.rs  # HTTP 请求上下文
    └── error_page.rs       # 错误页面
```

## 模块职责

| 目录 | 职责 | 状态 |
|------|------|------|
| `src/ast/` | 抽象语法树定义，纯数据结构 | ✅ 完成 |
| `src/parser/lexer.rs` | 词法分析，Token 流生成 | ✅ 完成 |
| `src/parser/expr_parser.rs` | 表达式解析（Pratt 算法）| ✅ 完成 |
| `src/parser/stmt_parser/` | 语句解析（递归下降）| ✅ 完成 |
| `src/runtime/` | 解释执行 AST，变量作用域 | ✅ 完成 |
| `src/runtime/value/` | Value 类型系统 | ✅ 完成 |
| `src/builtins/` | ASP 内建对象 | ⚠️ 部分完成 |
| `src/asp/engine.rs` | ASP 执行引擎 | ✅ 完成 |
| `src/asp/segmenter.rs` | 代码分段（HTML/代码分离）| ✅ 完成 |
| `src/asp/include.rs` | Include 指令处理 | ✅ 完成 |
| `src/http/` | Web 服务 | ✅ 完成 |

## 关键数据流

```
HTTP 请求
    ↓
http/handler.rs（加载 ASP 文件）
    ↓
asp/include.rs（处理 include 指令）
    ↓
asp/segmenter.rs（分割 HTML 与代码块）
    ↓
parser/lexer.rs（源代码 → Token 流）
    ↓
parser/stmt_parser/（Token → AST）
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
- 语句解析器拆分为多个文件：`stmt_parser/`

### 子集策略
支持：Dim, If, For, While, Function, Sub, Select Case, 基本表达式, Response.Write

不支持：COM, ActiveX, Windows-only DLL

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
- html-escape (HTML 转义)

## 注意事项

### 语句解析器架构

`stmt_parser/` 目录采用递归下降解析器，按功能拆分：
- `core.rs` - 核心解析逻辑和入口
- `control.rs` - 控制流语句（If, For, While, Do Loop, Select Case）
- `declarations.rs` - 声明语句（Dim, Const, ReDim）
- `procedures.rs` - 函数和过程（Function, Sub）
- `helpers.rs` - 辅助函数

### ASP 执行流程

1. `segmenter.rs` 将 ASP 文件分割为 HTML 段和代码段
2. `engine.rs` 对每个代码段执行：Lexer → Parser → Interpreter
3. 支持三种代码块：`<% %>` (代码), `<%= %>` (表达式), `<%@ %>` (指令)

### Include 指令

`include.rs` 支持：
- `<!--#include file="path" -->` - 相对于当前文件
- `<!--#include virtual="/path" -->` - 相对于网站根目录
- 自动检测循环 include

### 编译时间控制

如果发现编译时间过长：
- 检查是否有复杂的泛型嵌套
- 检查是否有过多的 derive 宏
- 考虑拆分大文件
- 使用 `cargo check` 代替 `cargo build` 进行快速验证
