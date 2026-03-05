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
ASP 引擎层 (segmenter + engine)
    ↓
VBScript Runtime (interpreter + value)
    ↓
Parser / AST (手写 Lexer + Pratt Parser)
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
│   └── error.rs            # 解析错误
│
├── runtime/                # 解释执行层
│   ├── mod.rs
│   ├── interpreter.rs      # 解释器调度
│   ├── context.rs          # 执行上下文
│   ├── scope.rs            # 变量作用域
│   ├── error.rs            # 运行时错误
│   └── value/              # 值类型系统
│       ├── mod.rs
│       ├── value.rs        # Value 定义
│       ├── conversion.rs   # 类型转换
│       ├── operators.rs    # 运算操作
│       ├── compare.rs      # 比较操作
│       └── display.rs      # 显示格式
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

### ast/ - 抽象语法树层
- 定义 VBScript 语法结构
- 只包含数据结构，不包含执行逻辑
- 支持静态分析、代码迁移工具

### parser/ - 语法解析层
- **lexer.rs**: 手写词法分析器，将源代码转为 Token 流
- **expr_parser.rs**: Pratt 算法表达式解析器
- ⚠️ **缺失**: 语句解析器（Statement Parser）

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
- 分割 HTML 与 `<% %>`
- 当前使用简化引擎（simple_engine.rs）
- 执行脚本片段，拼接输出

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
| **StmtParser** | 语句解析 | **递归下降解析器**，支持 Dim/If/For/While/Function/Sub |
| AST | 语法树定义 | 完整的语句和表达式定义 |
| Runtime | 解释执行 | 支持变量、If、For、While、Function 等 |
| Value | 类型系统 | 弱类型系统，类型转换，运算操作 |
| HTTP Server | Web 服务 | 支持 GET/POST，静态文件，ASP 执行 |
| RequestContext | 请求处理 | 解析 QueryString、Form 数据 |

### ⚠️ 进行中

| 模块 | 问题 | 说明 |
|------|------|------|
| SimpleEngine | 功能受限 | 仅支持简单的 Response.Write 和变量赋值 |
| 完整引擎集成 | 待完成 | 需要 StmtParser + Runtime 与 ASP 引擎集成 |

### ❌ 待实现

| 模块 | 说明 |
|------|------|
| Response 对象 | 完整实现（Buffer、ContentType、Redirect 等） |
| Session 管理 | 会话状态管理 |
| SQLite 支持 | Access 数据库迁移 |

## 历史背景

### 为什么有 SimpleEngine？

早期尝试实现完整的语句解析器（StatementParser），但由于：
- chumsky 解析器组合子编译时间过长（58分钟）
- 编译器内存占用巨大

因此采用**渐进式策略**：
1. 保留完整的 Lexer（快速编译）
2. 保留完整的 AST 定义
3. 保留完整的 Runtime
4. 使用 SimpleEngine 作为过渡方案
5. **渐进式接入语句解析**：一次只实现少量语句类型，确保每次都能快速编译

### 渐进式接入策略

```rust
// 第一步：实现 Dim 语句解析
// 第二步：实现 If 语句解析
// 第三步：实现 For 循环解析
// ...逐步扩展
```

每次只增加 1-2 种语句的解析，保证编译时间在可接受范围内。

## 使用流程

当前使用简化引擎的流程：

```rust
// 1. HTTP 请求到达 handler.rs
// 2. 构建 RequestContext（包含 QueryString、Form 数据）
let request_ctx = RequestContext::from_request(request).await;

// 3. 创建 SimpleEngine 并执行 ASP 文件
let mut engine = Engine::new()
    .with_debug(true)
    .with_request_context(request_ctx);
let output = engine.execute(&asp_content)?;

// 4. 返回 HTML 响应
```

完整引擎的使用流程（待实现语句解析器后）：

```rust
// 1. Lexer: 源代码 -> Token 流
let tokens = tokenize(source)?;

// 2. StatementParser: Token 流 -> AST (Program)  ← 缺失
let program = parse_program(tokens)?;

// 3. Interpreter: AST -> 执行
let mut interpreter = Interpreter::new();
interpreter.execute(&program)?;
```

## 设计原则

### 单一职责
每个目录代表一个系统维度：语法、执行、内建对象、模板引擎、HTTP

### 控制文件大小
- 单文件不超过 500 行
- match 分支过多必须拆文件
- Value 相关逻辑集中在 value/ 目录

### 子集策略
支持：Dim, If, Function, 基本表达式, Response.Write

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

## 技术栈

- Rust 2021 Edition
- axum 0.7 (Web 框架)
- tower / tower-http (中间件)
- tokio (异步运行时)
- serde (序列化)
- thiserror / anyhow (错误处理)
- urlencoding (URL 编解码)
- mime_guess (MIME 类型检测)

## 开发路线

### 第一阶段 ✅
- [x] AST 定义
- [x] Lexer 词法分析器
- [x] ExprParser 表达式解析器
- [x] Runtime 解释器框架
- [x] HTTP 服务器
- [x] RequestContext 请求上下文
- [x] 支持 GET/POST 请求

### 第二阶段 🚧
- [ ] StatementParser 语句解析器
- [ ] 完整的 ASP 引擎（替换 SimpleEngine）
- [ ] Response 对象完整实现
- [ ] Request 对象完整实现
- [ ] Server 对象

### 第三阶段
- [ ] Session 管理
- [ ] SQLite 支持
- [ ] 完整错误处理
- [ ] 容器化部署
- [ ] 生产环境优化

## 总体定位

> 一个安全可控的 Classic ASP 子集运行时
> 一个面向迁移的过渡引擎
> 一个容器友好的 IIS 替代方案

不是 100% 复刻 IIS，而是"工程可控的现代化替代实现"。
