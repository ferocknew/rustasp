# asp-lite - Rust 实现的 Classic ASP 运行时

一个用 Rust 实现的 Classic ASP（VBScript 子集）运行时，摆脱 IIS，支持容器化部署。

- **GitHub**: https://github.com/ferocknew/rustasp
- **Docker Hub**: https://hub.docker.com/r/jonahfu/vbscript

## 项目目标

- 使用 **Rust** 实现一个 Classic ASP（VBScript 子集）运行时
- 摆脱 IIS，支持容器化部署
- 支持 Nginx 反向代理
- 强调可维护性、可扩展性、模块解耦
- 不支持windows 编译（因为没有意义，windows 自带IIS + ASP）

## 不支持的特性

> ⚠️ **重要提示**：本项目是 VBScript 的**子集实现**，以下特性**不在支持范围内**：

| 类别 | 不支持的特性 | 说明                                       |
|------|-------------|------------------------------------------|
| **COM/ActiveX** | COM 组件、ActiveX 控件 | Windows 特有的组件技术                          |
| **数据库** | ADODB.Connection、ADODB.Recordset | 建议后续会更新支持使用 Sqlite、mysql、Linux 下的ODBC 驱动 |
| **Windows DLL** | 任何 Windows-only DLL | 跨平台限制                                    |
| **Server 方法** | Execute、Transfer、GetLastError | 计划未来实现                                   |
| **错误处理** | On Error Resume Next、Err 对象 | 计划未来实现                                   |
| **部分内置对象** | Application 对象 | 计划未来实现                                   |

如果您的项目依赖上述特性，需要进行迁移适配

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
├── main.rs                     # 入口
├── lib.rs                      # 库入口
│
├── ast/                        # 抽象语法树（纯数据结构）
│   ├── expr.rs                 # 表达式
│   ├── op.rs                   # 运算符定义
│   ├── program.rs              # 程序
│   └── stmt.rs                 # 语句
│
├── parser/                     # 语法解析层
│   ├── lexer/                  # 词法分析器（手写）
│   │   ├── mod.rs              # Lexer 核心
│   │   ├── token.rs            # Token 定义
│   │   └── number.rs           # 数字解析（十六进制/八进制）
│   ├── keyword.rs              # 关键字定义
│   ├── error.rs                # 解析错误
│   ├── parser.rs               # 解析器入口
│   ├── expr/                   # 表达式解析（Pratt 算法）
│   │   ├── pratt.rs            # Pratt 解析器核心
│   │   ├── prefix.rs           # 前缀表达式
│   │   ├── infix.rs            # 中缀表达式
│   │   └── postfix.rs          # 后缀表达式
│   └── stmt/                   # 语句解析（递归下降）
│       ├── core.rs             # 核心解析逻辑
│       ├── assign_stmt.rs      # 赋值语句
│       ├── decl_stmt.rs        # 声明语句（Dim/Const/ReDim）
│       ├── if_stmt.rs          # If 条件语句
│       ├── loop_stmt.rs        # 循环语句（For/While/Do/For Each）
│       ├── proc_stmt.rs        # 函数/过程/Class
│       ├── select_stmt.rs      # Select Case 语句
│       └── with_stmt.rs        # With 语句
│
├── runtime/                    # 解释执行层
│   ├── interpreter/            # 解释器调度
│   ├── context.rs              # 执行上下文
│   ├── scope.rs                # 变量作用域
│   ├── class.rs                # Class 运行时支持
│   ├── error.rs                # 运行时错误
│   ├── value/                  # 值类型系统
│   ├── objects/                # ASP 内建对象
│   │   ├── response.rs         # Response 对象
│   │   ├── request.rs          # Request 对象
│   │   ├── server.rs           # Server 对象
│   │   ├── session.rs          # Session 对象
│   │   ├── session_manager.rs  # Session 管理
│   │   ├── session_store.rs    # Session 存储
│   │   ├── dictionary.rs       # Scripting.Dictionary
│   │   ├── filesystemobject.rs # Scripting.FileSystemObject
│   │   ├── xmlhttp.rs          # MSXML2.XMLHTTP
│   │   └── factory.rs          # 对象工厂
│   └── builtins/               # 内置函数
│       ├── token.rs            # 函数 Token ID
│       ├── registry.rs         # 函数注册表
│       └── executors.rs        # 函数执行器
│
├── asp/                        # ASP 引擎层
│   ├── engine/                 # 引擎子模块
│   ├── segmenter.rs            # 代码分段器
│   └── include.rs              # Include 指令处理
│
└── http/                       # HTTP 服务层
    ├── router.rs               # 路由配置
    ├── handler.rs              # 请求处理
    ├── state.rs                # 应用状态
    ├── path_resolver.rs        # 路径解析（安全防护）
    ├── request_context.rs      # HTTP 请求上下文
    └── error_page.rs           # 错误页面
```

## 模块职责

### ast/ - 抽象语法树层
- 定义 VBScript 语法结构
- 只包含数据结构，不包含执行逻辑
- 支持静态分析、代码迁移工具

### parser/ - 语法解析层
- **lexer/**: 手写词法分析器，将源代码转为 Token 流，支持十六进制/八进制数字
- **expr/**: Pratt 算法表达式解析器（prefix/infix/postfix）
- **stmt/**: 递归下降语句解析器，支持所有 VBScript 语句类型（含 Class/With）

### runtime/ - 解释执行层
- **interpreter/**: 执行 AST 语句
- 管理变量作用域和函数调用
- **class.rs**: VBScript Class 运行时支持
- 实现 VBScript 弱类型系统
- **value/**: 值类型系统，隔离弱类型逻辑
- **objects/**: ASP 内建对象（Response/Request/Server/Session/Dictionary/FileSystemObject/XMLHTTP）
- **builtins/**: 内置函数注册和执行

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
| Lexer | 词法分析 | 手写实现，支持关键字、标识符、字符串、数字（含十六进制/八进制）、运算符 |
| ExprParser | 表达式解析 | Pratt 算法，支持算术、逻辑、比较运算 |
| StmtParser | 语句解析 | 递归下降解析器，支持 Dim/If/For/While/Function/Sub/Select Case/Class/With |
| AST | 语法树定义 | 完整的语句和表达式定义 |
| Runtime | 解释执行 | 支持变量、If、For、While、Function、Sub、Class、With 等 |
| Value | 类型系统 | 弱类型系统，类型转换，运算操作 |
| ASP Engine | ASP 引擎 | 完整集成 Lexer → Parser → Interpreter |
| HTTP Server | Web 服务 | 支持 GET/POST，静态文件，ASP 执行 |
| Session | 会话管理 | 支持 memory/json/redis 三种存储模式 |
| Include | Include 指令 | 支持 file/virtual 路径，自动检测循环引用 |
| Class | 类支持 | VBScript Class 定义、实例化、方法调用、属性访问 |
| With | With 语句 | 简化对象成员访问 |
| Server.CreateObject | 对象创建 | 支持 Dictionary、FileSystemObject、XMLHTTP |

### ⚠️ 部分完成

| 模块 | 功能 | 说明 |
|------|------|------|
| Response | HTTP 响应 | 基础功能完成，Buffer、ContentType、Redirect 待完善 |
| Request | HTTP 请求 | QueryString、Form 数据解析完成 |
| Server | 服务端方法 | MapPath、URLEncode、HTMLEncode、ScriptTimeout 已实现<br>**暂未支持**: Execute、Transfer、GetLastError |

### ❌ 待实现

| 模块 | 说明 |
|------|------|
| 数据库支持 | SQLite/MySQL/ODBC 数据库访问（计划中） |
| 错误处理 | On Error Resume Next、Err 对象 |
| Application 对象 | 全局应用程序对象 |
| Server 方法 | Execute、Transfer、GetLastError |
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

**✅ 已支持：**
- 变量声明：Dim, Const, ReDim
- 流程控制：If/Then/Else, Select Case, For/Next, For Each/Next, While/Wend, Do/Loop
- 过程：Function, Sub, Exit Function/Sub
- 类：Class, Property Get/Let/Set
- 语句：With, Call, Erase
- 表达式：算术、逻辑、比较、字符串连接
- 内置对象：Response, Request, Server, Session, Dictionary, FileSystemObject, XMLHTTP
- ASP 特性：Include 指令, Session 管理

**❌ 不支持：**
- COM/ActiveX 组件
- ADODB 数据库访问
- Windows-only DLL
- On Error Resume Next 错误处理
- Application 对象
- Server.Execute/Transfer/GetLastError

## 快速开始

### Docker 部署（推荐）

```bash
# 拉取镜像
docker pull jonahfu/vbscript:latest

# 运行容器
docker run -d \
  --name vbscript \
  -p 8080:8080 \
  -v $(pwd)/www:/app/www \
  jonahfu/vbscript:latest
```

### Docker Compose 部署

创建 `docker-compose.yml` 文件：

```yaml
version: '3.8'

services:
  vbscript:
    image: jonahfu/vbscript:latest
    container_name: vbscript
    restart: unless-stopped
    ports:
      - "8080:8080"
    volumes:
      # 挂载 ASP 源码目录
      - ./www:/app/www
      # 挂载 Session 持久化目录（可选）
      - ./runtime/sessions:/app/runtime/sessions
    environment:
      # ========== 基础 HTTP 服务配置 ==========
      # 是否显示目录列表
      - DIRECTORY_LISTING=false
      # Web 根目录路径（容器内路径）
      - HOME_DIR=/app/www
      # 默认索引文件
      - INDEX_FILE=index.asp,index.html
      # 索引文件功能开关
      - INDEX_FILE_ENABLE=true
      # 服务端口号
      - PORT=8080
      # 是否支持父路径访问（安全风险）
      - ALLOW_PARENT_PATH=false
      # ASP 文件扩展名
      - ASP_EXT=asp,asa

      # ========== 调试与错误处理配置 ==========
      # 是否启用调试模式
      - DEBUG=false
      # 是否显示详细错误信息（生产环境建议关闭）
      - DETAILED_ERROR=false
      # 自定义错误页面（相对于 HOME_DIR）
      # - ERROR_PAGE=error.html

      # ========== Session 会话管理配置 ==========
      # Session 存储模式 (memory/json/redis)
      # memory - 纯内存存储，重启后丢失
      # json - JSON文件持久化存储
      # redis - Redis存储（预留）
      - SESSION_STORAGE=memory
      # Session 超时时间（分钟）
      - SESSION_TIMEOUT=20
      # Session 存储目录（仅 json 模式）
      - SESSION_DIR=/app/runtime/sessions
      # Redis 连接配置（预留，仅 redis 模式）
      # - REDIS_URL=redis://redis:6379
      # - REDIS_KEY_PREFIX=vbscript:session:

      # ========== 日期时间格式配置 ==========
      # 格式说明：yyyy=年 mm=月 dd=日 hh=时 nn=分 ss=秒
      - NOW_FORMAT=yyyy/mm/dd hh:nn:ss
      - DATE_FORMAT=yyyy/mm/dd
      - TIME_FORMAT=hh:nn:ss

      # ========== CreateObject 安全配置 ==========
      # CreateObject 功能开关
      - CREATE_OBJECT_ENABLE=true
      # Server.CreateObject 白名单
      - CREATE_OBJECT_WHITELIST=Scripting.Dictionary,Scripting.FileSystemObject,MSXML2.XMLHTTP

  # 可选：Redis 服务（用于 Session 持久化，redis 驱动还未支持，后续会支持）
  # redis:
  #   image: redis:7-alpine
  #   container_name: vbscript-redis
  #   restart: unless-stopped
  #   volumes:
  #     - redis_data:/data

# volumes:
#   redis_data:
```

启动服务：

```bash
# 启动
docker-compose up -d

# 查看日志
docker-compose logs -f

# 停止
docker-compose down
```

### 配置

在项目根目录创建 `.env` 文件：

```env
# ==================== 基础配置 ====================

# 是否显示目录列表 (true/false)
DIRECTORY_LISTING=false

# Web 根目录路径
HOME_DIR=./www

# 默认索引文件（支持多个，逗号分隔）
INDEX_FILE=index.asp,index.html

# 索引文件功能开关 (true/false)
INDEX_FILE_ENABLE=true

# 服务端口号
PORT=8080

# 是否支持父路径访问 (../) (true/false)
# 警告：启用此选项可能存在安全风险
ALLOW_PARENT_PATH=false

# ASP 文件扩展名（逗号分隔）
ASP_EXT=asp,asa

# ==================== 调试配置 ====================

# 是否启用调试模式 (true/false)
DEBUG=false

# 是否显示详细错误信息 (true/false)
# 生产环境建议关闭
DETAILED_ERROR=false

# 自定义错误页面（相对于 home_dir）
# ERROR_PAGE=error.html

# ==================== Session 配置 ====================

# Session 存储模式 (memory/json/redis)
# memory - 纯内存存储，重启后丢失
# json - JSON文件持久化存储（默认）
# redis - Redis存储（预留接口）
SESSION_STORAGE=memory

# Session 超时时间（分钟）
SESSION_TIMEOUT=20

# Session 存储目录（仅 json 模式使用）
SESSION_DIR=./runtime/sessions

# Redis 连接地址（预留，仅 redis 模式使用）
# REDIS_URL=redis://127.0.0.1:6379
# REDIS_KEY_PREFIX=vbscript:session:

# ==================== 日期时间格式配置 ====================
# 参考：https://learn.microsoft.com/zh-tw/office/vba/language/reference/user-interface-help/format-function-visual-basic-for-applications
#
# 格式说明：
#   yyyy = 四位年份 (2024)
#   yy = 两位年份 (24)
#   mm = 月份，带前导零 (01-12)
#   m = 月份，不带前导零 (1-12)
#   dd = 日期，带前导零 (01-31)
#   d = 日期，不带前导零 (1-31)
#   hh = 小时，带前导零 (00-23)
#   h = 小时，不带前导零 (0-23)
#   nn = 分钟，带前导零 (00-59)
#   n = 分钟，不带前导零 (0-59)
#   ss = 秒，带前导零 (00-59)
#   s = 秒，不带前导零 (0-59)
NOW_FORMAT=yyyy/mm/dd hh:nn:ss
DATE_FORMAT=yyyy/mm/dd
TIME_FORMAT=hh:nn:ss

# ==================== CreateObject 配置 ====================

# CreateObject 功能开关 (true/false)
CREATE_OBJECT_ENABLE=true

# Server.CreateObject 白名单（逗号分隔）
# 出于安全考虑，只允许创建以下对象：
# - Scripting.Dictionary: 字典对象，用于键值对存储
# - Scripting.FileSystemObject: 文件系统对象，用于文件和文件夹操作
# - MSXML2.XMLHTTP: HTTP 请求对象，用于发起 HTTP 请求
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
