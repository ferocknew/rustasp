# asp-lite - Rust 实现的 Classic ASP 运行时

一个用 Rust 实现的 Classic ASP（VBScript 子集）运行时，摆脱 IIS，支持容器化部署。

## 项目目标

- 使用 **Rust** 实现一个 Classic ASP（VBScript 子集）运行时
- 摆脱 IIS，支持容器化部署
- 不支持 COM / ActiveX / Windows-only DLL
- Access 迁移为 SQLite
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
Parser / AST (chumsky)
```

## 目录结构

```
src/
├── main.rs              # 入口
│
├── ast/                 # 抽象语法树（纯数据结构）
│   ├── mod.rs
│   ├── op.rs            # 运算符定义
│   ├── expr.rs          # 表达式
│   ├── stmt.rs          # 语句
│   └── program.rs       # 程序
│
├── parser/              # 语法解析层
│   ├── mod.rs
│   ├── keyword.rs       # 关键字定义
│   ├── lexer.rs         # 词法分析
│   ├── parser.rs        # 语法分析
│   └── error.rs         # 解析错误
│
├── runtime/             # 解释执行层
│   ├── mod.rs
│   ├── interpreter.rs   # 解释器调度
│   ├── context.rs       # 执行上下文
│   ├── scope.rs         # 变量作用域
│   ├── error.rs         # 运行时错误
│   └── value/           # 值类型系统
│       ├── mod.rs
│       ├── value.rs     # Value 定义
│       ├── conversion.rs # 类型转换
│       ├── operators.rs  # 运算操作
│       ├── compare.rs    # 比较操作
│       └── display.rs    # 显示格式
│
├── builtins/            # ASP 内建对象
│   ├── mod.rs
│   ├── response.rs      # Response 对象
│   ├── request.rs       # Request 对象
│   ├── server.rs        # Server 对象
│   └── session.rs       # Session 对象
│
├── asp/                 # ASP 引擎层
│   ├── mod.rs
│   ├── engine.rs        # 执行引擎
│   └── segmenter.rs     # 代码分段
│
└── http/                # HTTP 服务层
    ├── mod.rs
    ├── router.rs        # 路由配置
    ├── handler.rs       # 请求处理
    └── state.rs         # 应用状态
```

## 模块职责

### ast/ - 抽象语法树层
- 定义 VBScript 语法结构
- 只包含数据结构，不包含执行逻辑
- 支持静态分析、代码迁移工具

### parser/ - 语法解析层
- 词法分析（lexer）
- 语法分析（parser）
- 输出 AST，不执行代码

### runtime/ - 解释执行层
- 执行 AST
- 管理变量作用域和函数调用
- 实现 VBScript 弱类型系统
- value/ 子目录隔离弱类型逻辑

### builtins/ - ASP 内建对象层
- Response, Request, Server, Session
- 只实现白名单功能
- 不支持 COM/ActiveX

### asp/ - ASP 引擎层
- 分割 HTML 与 `<% %>`
- 执行脚本片段
- 拼接输出

### http/ - Web 服务层
- 路由和文件加载
- HTTP 响应
- 不处理 VBScript 逻辑

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
DIRECTORY_LISTING=true

# Web 根目录路径
HOME_DIR=./www

# 默认索引文件
INDEX_FILE=index.asp

# 服务端口号
PORT=8080
```

### 运行

```bash
# 开发模式
cargo run

# Release 模式
cargo run --release
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
- axum 0.7 - Web 框架
- tower - 中间件
- chumsky 0.9 - 解析器组合子
- tokio - 异步运行时
- serde - 序列化

## 开发阶段

### 第一阶段（当前）
- [x] AST 定义
- [x] Parser (lexer + parser)
- [x] 基本 Runtime
- [x] HTTP 服务器
- [ ] `<%= %>` 表达式输出
- [ ] Response.Write

### 第二阶段
- [ ] 函数支持
- [ ] 作用域管理
- [ ] SQLite 支持
- [ ] Request 对象

### 第三阶段
- [ ] Session 管理
- [ ] 完整错误处理
- [ ] 容器化部署
- [ ] 生产环境优化

## 总体定位

> 一个安全可控的 Classic ASP 子集运行时
> 一个面向迁移的过渡引擎
> 一个容器友好的 IIS 替代方案

不是 100% 复刻 IIS，而是"工程可控的现代化替代实现"。
