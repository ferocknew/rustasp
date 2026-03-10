# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

Rust 实现的 Classic ASP（VBScript 子集）运行时，摆脱 IIS，支持容器化部署。

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

# 交叉编译 Linux x64
cargo install cross
cross build --release --target x86_64-unknown-linux-musl
```

## 架构分层

```
HTTP 层 (axum)
    ↓
ASP 引擎层 (segmenter + engine + include)
    ↓
VBScript Runtime (interpreter + value + objects)
    ↓
Parser / AST (手写 Lexer + Pratt 表达式解析器 + 递归下降语句解析器)
```

## 目录结构

```
src/
├── main.rs                     # 入口
├── lib.rs                      # 库入口
├── ast/                        # 抽象语法树（纯数据结构）
├── parser/                     # 语法解析层
│   ├── lexer/                  # 词法分析器（手写）
│   ├── expr/                   # 表达式解析（Pratt 算法）
│   └── stmt/                   # 语句解析（递归下降）
├── runtime/                    # 解释执行层
│   ├── interpreter/            # 解释器调度
│   ├── context.rs              # 执行上下文
│   ├── scope.rs                # 变量作用域
│   ├── class.rs                # Class 支持
│   ├── value/                  # 值类型系统
│   ├── objects/                # ASP 内建对象
│   └── builtins/               # 内置函数
├── asp/                        # ASP 引擎层
│   ├── engine/                 # 引擎子模块
│   ├── segmenter.rs            # 代码分段器
│   └── include.rs              # Include 指令处理
└── http/                       # HTTP 服务层
```

## 关键数据流

```
HTTP 请求 → http/handler.rs → asp/include.rs → asp/segmenter.rs
    → parser/lexer → parser/stmt → runtime/interpreter
    → runtime/objects → HTML 输出
```

## 设计原则

- **单文件不超过 500 行**：match 分支过多必须拆文件
- **子集策略**：支持 Dim/If/For/While/Do Loop/Function/Sub/Select Case/Class/With
- **不支持**：COM, ActiveX, Windows-only DLL

## 配置（.env 文件）

```env
HOME_DIR=./www              # Web 根目录
PORT=8080                   # 服务端口
DEBUG=false                 # 调试模式
DETAILED_ERROR=false        # 详细错误信息
SESSION_STORAGE=memory      # Session 存储 (memory/json/redis)
SESSION_TIMEOUT=20          # Session 超时（分钟）
CREATE_OBJECT_WHITELIST=... # Server.CreateObject 白名单
```

## 内置对象（runtime/objects/）

| 对象 | 文件 | 说明 |
|------|------|------|
| Response | response.rs | Write/End/Redirect/Buffer |
| Request | request.rs | QueryString/Form/ServerVariables |
| Server | server.rs | MapPath/HTMLEncode/URLEncode/CreateObject |
| Session | session.rs | 会话管理 |
| Dictionary | dictionary.rs | Scripting.Dictionary 实现 |
| FileSystemObject | filesystemobject.rs | 文件系统操作 |
| XMLHTTP | xmlhttp.rs | MSXML2.XMLHTTP 实现 |

## 解析器架构

### 表达式（Pratt 算法）
- `parser/expr/pratt.rs` - 核心
- `parser/expr/prefix.rs` - 前缀（字面量、一元运算、函数调用）
- `parser/expr/infix.rs` - 中缀（二元运算）
- `parser/expr/postfix.rs` - 后缀（成员访问、数组索引）

### 语句（递归下降）
- `parser/stmt/core.rs` - 核心入口
- `parser/stmt/assign_stmt.rs` - 赋值语句
- `parser/stmt/decl_stmt.rs` - Dim/Const/ReDim
- `parser/stmt/if_stmt.rs` - If 条件
- `parser/stmt/loop_stmt.rs` - For/While/Do Loop
- `parser/stmt/proc_stmt.rs` - Function/Sub
- `parser/stmt/select_stmt.rs` - Select Case
- `parser/stmt/with_stmt.rs` - With 语句

## BuiltinObject Trait

所有内置对象实现 `BuiltinObject` trait：

```rust
pub trait BuiltinObject: Send + Sync + std::fmt::Debug {
    fn get_property(&self, name: &str) -> Result<Value, RuntimeError>;
    fn set_property(&mut self, name: &str, value: Value) -> Result<(), RuntimeError>;
    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError>;
    fn index(&self, key: &Value) -> Result<Value, RuntimeError>;  // Session("key")
    fn set_index(&mut self, key: &Value, value: Value) -> Result<(), RuntimeError>;
}
```

## ASP 执行流程

1. `segmenter.rs` 将 ASP 文件分割为 HTML 段和代码段
2. `engine.rs` 对每个代码段执行：Lexer → Parser → Interpreter
3. 支持三种代码块：`<% %>` / `<%= %>` / `<%@ %>`
