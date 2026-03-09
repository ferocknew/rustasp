# asp-lite - Rust 实现的 Classic ASP 运行时

一个用 Rust 实现的 Classic ASP（VBScript 子集）运行时，摆脱 IIS，支持容器化部署。

## 项目目标

- 使用 **Rust** 实现一个 Classic ASP（VBScript 子集）运行时
- 摆脱 IIS，支持容器化部署
- 不支持 COM / ActiveX / Windows-only DLL
- 支持 Nginx 反向代理
- 强调可维护性、可扩展性、模块解耦

## 架构分层

```
HTTP 层 (axum)
    ↓
ASP 引擎层 (segmenter + engine + include)
    ↓
VBScript Runtime (interpreter + value)
    ↓
Parser / AST (手写 Lexer + Pratt 表达式解析器 + 递归下降语句解析器)
```

## 目录结构

```
src/
├── main.rs                 # 入口
├── lib.rs                  # 库入口
│
├── ast/                    # 抽象语法树（纯数据结构）
│   ├── expr.rs             # 表达式
│   ├── op.rs               # 运算符定义
│   ├── program.rs          # 程序
│   └── stmt.rs             # 语句
│
├── parser/                 # 语法解析层
│   ├── lexer.rs            # 词法分析器（手写）
│   ├── keyword.rs          # 关键字定义
│   ├── error.rs            # 解析错误
│   ├── parser.rs           # 解析器入口
│   ├── program.rs          # 程序解析
│   ├── expr/               # 表达式解析（Pratt 算法）
│   │   ├── pratt.rs        # Pratt 解析器核心
│   │   ├── prefix.rs       # 前缀表达式
│   │   ├── infix.rs        # 中缀表达式
│   │   └── postfix.rs      # 后缀表达式
│   └── stmt/               # 语句解析（递归下降）
│       ├── core.rs         # 核心解析逻辑
│       ├── assign_stmt.rs  # 赋值语句
│       ├── decl_stmt.rs    # 声明语句（Dim/Const/ReDim）
│       ├── if_stmt.rs      # If 条件语句
│       ├── loop_stmt.rs    # 循环语句（For/While/Do）
│       ├── proc_stmt.rs    # 函数/过程
│       └── select_stmt.rs  # Select Case 语句
│
├── runtime/                # 解释执行层
│   ├── interpreter.rs      # 解释器调度
│   ├── context.rs          # 执行上下文
│   ├── scope.rs            # 变量作用域
│   ├── error.rs            # 运行时错误
│   └── value/              # 值类型系统
│
├── builtins/               # ASP 内建对象
│   ├── response.rs         # Response 对象
│   ├── request.rs          # Request 对象
│   ├── server.rs           # Server 对象
│   └── session.rs          # Session 对象（memory/json/redis）
│
├── asp/                    # ASP 引擎层
│   ├── engine.rs           # ASP 执行引擎
│   ├── engine/             # 引擎子模块
│   ├── segmenter.rs        # 代码分段器
│   └── include.rs          # Include 指令处理
│
└── http/                   # HTTP 服务层
    ├── router.rs           # 路由配置
    ├── handler.rs          # 请求处理
    ├── state.rs            # 应用状态
    ├── path_resolver.rs    # 路径解析（安全防护）
    ├── request_context.rs  # HTTP 请求上下文
    └── error_page.rs       # 错误页面
```

## 模块职责

### ast/ - 抽象语法树层
- 定义 VBScript 语法结构
- 只包含数据结构，不包含执行逻辑
- 支持静态分析、代码迁移工具

### parser/ - 语法解析层
- **lexer.rs**: 手写词法分析器，将源代码转为 Token 流
- **expr/**: Pratt 算法表达式解析器（prefix/infix/postfix）
- **stmt/**: 递归下降语句解析器，支持所有 VBScript 语句类型

### runtime/ - 解释执行层
- 执行 AST 语句
- 管理变量作用域和函数调用
- 实现 VBScript 弱类型系统
- value/ 子目录隔离弱类型逻辑

### builtins/ - ASP 内建对象层
- Response, Request, Server, Session
- 只实现白名单功能
- 不支持 COM/ActiveX

### asp/ - ASP 引擎层
- **segmenter.rs**: 分割 HTML 与 `<% %>` 代码块
- **engine.rs**: ASP 执行引擎，集成 Lexer → Parser → Interpreter
- **include.rs**: 支持 file/virtual 两种路径的 Include 指令
- 支持三种代码块：`<% %>` (代码), `<%= %>` (表达式), `<%@ %>` (指令)

### http/ - Web 服务层
- 路由和文件加载
- HTTP 响应处理
- 请求上下文构建（支持 GET/POST）
- 路径安全解析（防止目录遍历攻击）

## 当前实现状态

### ✅ 已完成

| 模块 | 功能 | 说明 |
|------|------|------|
| Lexer | 词法分析 | 手写实现，支持关键字、标识符、字符串、数字、运算符 |
| ExprParser | 表达式解析 | Pratt 算法，支持算术、逻辑、比较运算 |
| StmtParser | 语句解析 | 递归下降解析器，支持 Dim/If/For/While/Function/Sub/Select Case |
| AST | 语法树定义 | 完整的语句和表达式定义 |
| Runtime | 解释执行 | 支持变量、If、For、While、Function、Sub 等 |
| Value | 类型系统 | 弱类型系统，类型转换，运算操作 |
| ASP Engine | ASP 引擎 | 完整集成 Lexer → Parser → Interpreter |
| HTTP Server | Web 服务 | 支持 GET/POST，静态文件，ASP 执行 |
| Session | 会话管理 | 支持 memory/json/redis 三种存储模式 |
| Include | Include 指令 | 支持 file/virtual 路径，自动检测循环引用 |

### ⚠️ 部分完成

| 模块 | 功能 | 说明 |
|------|------|------|
| Response | HTTP 响应 | 基础功能完成，Buffer、ContentType、Redirect 待完善 |
| Request | HTTP 请求 | QueryString、Form 数据解析完成 |
| Server | 服务端方法 | MapPath、URLEncode、HTMLEncode、ScriptTimeout 已实现<br>**暂未支持**: Execute、Transfer、GetLastError |

### ❌ 待实现

| 模块 | 说明 |
|------|------|
| 数据库支持 | SQLite/Access 数据库访问 |
| 完整错误处理 | 详细的错误信息和调试支持 |
| 容器化部署 | Docker 镜像和部署文档 |

## 执行流程

完整的 ASP 请求处理流程：

```rust
// 1. HTTP 请求到达 handler.rs
// 2. 加载 ASP 文件，处理 include 指令
let asp_content = include_processor.resolve_includes(file_path)?;

// 3. 分段器将 ASP 文件分割为 HTML 段和代码段
let segments = segmenter.segment(&asp_content);

// 4. 对每个代码段执行：Lexer → Parser → Interpreter
for segment in segments {
    match segment {
        Segment::Html(html) => output.push_str(html),
        Segment::Code(code) => {
            let tokens = lexer.tokenize(&code)?;
            let program = parser.parse_program(tokens)?;
            interpreter.execute(&program)?;
        }
    }
}

// 5. 返回 HTML 响应
```

## 设计原则

### 单一职责
每个目录代表一个系统维度：语法、执行、内建对象、模板引擎、HTTP

### 控制文件大小
- **单文件不超过 500 行**
- match 分支过多必须拆文件
- Value 相关逻辑集中在 `runtime/value/` 目录
- 语句解析器拆分为多个文件：`parser/stmt/`
- 表达式解析器拆分为多个文件：`parser/expr/`

### 子集策略
支持：Dim, If, For, While, Do Loop, Function, Sub, Select Case, 基本表达式, Response.Write, Session 管理

不支持：COM, ActiveX, Windows-only DLL

## 快速开始

### 配置

在项目根目录创建 `.env` 文件：

```env
# 是否显示目录列表 (true/false)
DIRECTORY_LISTING=false

# Web 根目录路径
HOME_DIR=./www

# 默认索引文件
INDEX_FILE=index.asp

# 服务端口号
PORT=8080

# 是否支持父路径访问 (../) (true/false)
ALLOW_PARENT_PATH=false

# 是否启用调试模式 (true/false)
DEBUG=false

# 是否显示详细错误信息 (true/false)
DETAILED_ERROR=false

# 自定义错误页面（相对于 home_dir）
# ERROR_PAGE=error.html

# ASP 文件扩展名（逗号分隔）
ASP_EXT=asp,asa

# CreateObject 功能配置
CREATE_OBJECT_ENABLE=true
# Server.CreateObject 白名单（逗号分隔）
# 出于安全考虑，只支持以下 3 个对象：
# - Scripting.Dictionary: 字典对象
# - Scripting.FileSystemObject: 文件系统对象
# - MSXML2.XMLHTTP: HTTP 请求对象
# 注意：不支持的 COM 对象（如 ADODB.Connection）将被拒绝
CREATE_OBJECT_WHITELIST=Scripting.Dictionary,Scripting.FileSystemObject,MSXML2.XMLHTTP
```

### 运行

```bash
# 开发模式
cargo run

# Release 模式
cargo run --release

# 带调试日志
DEBUG=true cargo run
```

服务器启动后会显示：
```
🚀 VBScript ASP Server starting at http://127.0.0.1:8080
📁 Home directory: ./www
📄 Index file: index.asp
📋 Directory listing: true
```

## 测试环境

### IIS 对比测试服务器

用于对比 VBScript 解析结果和 IIS 行为：

- **地址**: http://10.0.0.217/www/rust_vbs_test/
- **路径**: `/Volumes/Users/Administrator/Documents/www/rust_vbs_test`

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

## 开发路线

### 第一阶段 ✅
- [x] AST 定义
- [x] Lexer 词法分析器
- [x] ExprParser 表达式解析器（Pratt 算法）
- [x] StmtParser 语句解析器（递归下降）
- [x] Runtime 解释器框架
- [x] HTTP 服务器
- [x] RequestContext 请求上下文
- [x] 支持 GET/POST 请求

### 第二阶段 ✅
- [x] 完整的 ASP 引擎（Lexer → Parser → Interpreter 集成）
- [x] 代码分段器（HTML/代码分离）
- [x] Include 指令处理（file/virtual 路径）
- [x] Session 管理（memory/json/redis）
- [x] 核心内建对象（Response/Request/Server/Session）

### 第三阶段 🚧
- [ ] Response 对象完整功能（Buffer、ContentType、Redirect、End 等）
- [ ] 数据库支持（SQLite/Access 迁移）
- [ ] 完整错误处理和调试信息
- [ ] 性能优化和缓存
- [ ] 容器化部署
- [ ] 生产环境测试和优化

## 总体定位

> 一个安全可控的 Classic ASP 子集运行时
> 一个面向迁移的过渡引擎
> 一个容器友好的 IIS 替代方案

不是 100% 复刻 IIS，而是"工程可控的现代化替代实现"。
