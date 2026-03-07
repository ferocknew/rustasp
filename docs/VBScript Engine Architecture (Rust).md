# VBScript Engine Architecture (Rust)

## 1. Overall Architecture

```
HTTP Request
     │
     ▼
 ASP Parser (HTML + VBScript)
     │
     ▼
   VBScript Code
     │
     ▼
      Lexer
     │
     ▼
      Parser
     │
     ▼
       AST
     │
     ▼
   Interpreter
     │
     ├── Builtin Functions
     ├── Runtime Objects
     ├── Variables Scope
     │
     ▼
 Execution Result
     │
     ▼
HTTP Response
```

------

# 2. Project Directory Structure

推荐项目结构：

```
src/
│
├── engine/
│   ├── lexer/
│   │   ├── mod.rs
│   │   └── lexer.rs
│   │
│   ├── parser/
│   │   ├── mod.rs
│   │   ├── parser.rs
│   │   └── ast.rs
│   │
│   ├── interpreter/
│   │   ├── mod.rs
│   │   ├── interpreter.rs
│   │   ├── scope.rs
│   │   └── value.rs
│   │
│   └── runtime/
│       ├── mod.rs
│       │
│       ├── builtins/
│       │   ├── mod.rs
│       │   ├── string.rs
│       │   ├── math.rs
│       │   ├── date.rs
│       │   ├── array.rs
│       │   ├── convert.rs
│       │   └── random.rs
│       │
│       ├── objects/
│       │   ├── mod.rs
│       │   ├── request.rs
│       │   ├── response.rs
│       │   ├── session.rs
│       │   └── server.rs
│       │
│       └── helpers/
│           ├── mod.rs
│           └── convert.rs
│
├── asp/
│   ├── mod.rs
│   └── asp_parser.rs
│
└── main.rs
```

------

# 3. Core Components

## 3.1 Lexer

负责把 VBScript 代码转换为 Token。

输入：

```
a = 1 + 2
```

输出：

```
Identifier(a)
Equal
Number(1)
Plus
Number(2)
```

文件：

```
engine/lexer/lexer.rs
```

结构：

```
pub struct Lexer {
    input: String,
    position: usize,
}
```

主要方法：

```
pub fn next_token(&mut self) -> Token
```

------

# 3.2 Parser

Parser 把 Token 转换为 AST。

输入：

```
a = 1 + 2
```

输出：

```
Assignment
 ├── Identifier(a)
 └── BinaryExpr
     ├── Number(1)
     ├── +
     └── Number(2)
```

文件：

```
engine/parser/parser.rs
engine/parser/ast.rs
```

AST 示例：

```
pub enum Expr {
    Number(f64),
    String(String),
    Identifier(String),

    Binary {
        left: Box<Expr>,
        op: BinaryOp,
        right: Box<Expr>,
    },

    Call {
        name: String,
        args: Vec<Expr>,
    },
}
```

------

# 4. Interpreter

解释 AST 并执行代码。

文件：

```
engine/interpreter/interpreter.rs
```

核心结构：

```
pub struct Interpreter {
    scope: Scope,
}
```

执行入口：

```
pub fn eval(&mut self, expr: &Expr) -> Result<Value>
```

流程：

```
AST Node
   │
   ▼
Interpreter::eval
   │
   ├── 变量读取
   ├── 函数调用
   ├── 运算执行
   │
   ▼
Value
```

------

# 5. Value System

VBScript 的运行时值类型。

文件：

```
engine/interpreter/value.rs
```

定义：

```
pub enum Value {
    Null,
    Empty,
    Boolean(bool),
    Number(f64),
    String(String),

    Array(Vec<Value>),
    Object(Box<dyn RuntimeObject>)
}
```

------

# 6. Scope System

变量作用域管理。

文件：

```
engine/interpreter/scope.rs
```

结构：

```
pub struct Scope {
    variables: HashMap<String, Value>,
}
```

方法：

```
pub fn get(&self, name: &str) -> Option<Value>

pub fn set(&mut self, name: String, value: Value)
```

------

# 7. Builtin Functions

VBScript 内置函数。

目录：

```
runtime/builtins/
```

分类：

| 文件       | 功能       |
| ---------- | ---------- |
| string.rs  | 字符串函数 |
| math.rs    | 数学函数   |
| date.rs    | 日期函数   |
| array.rs   | 数组函数   |
| convert.rs | 类型转换   |
| random.rs  | 随机函数   |

例子：

```
pub fn instr(args: &[Value]) -> Result<Value>
```

注册表：

```
type BuiltinFn = fn(&[Value]) -> Result<Value>;
```

函数表：

```
static BUILTINS: phf::Map<&'static str, BuiltinFn>
```

调用：

```
if let Some(func) = BUILTINS.get(name) {
    return func(args);
}
```

------

# 8. Runtime Objects

ASP 对象系统。

目录：

```
runtime/objects/
```

对象：

| 对象     | 文件        |
| -------- | ----------- |
| Request  | request.rs  |
| Response | response.rs |
| Session  | session.rs  |
| Server   | server.rs   |

示例：

```
pub struct Response {
    buffer: String
}
```

方法：

```
pub fn write(&mut self, value: Value)
```

VBScript 调用：

```
Response.Write "Hello"
```

------

# 9. ASP Parser

ASP 页面解析。

输入：

```
Hello
<%
Response.Write "World"
%>
```

输出：

```
HTML("Hello")

VBScript(
    Response.Write "World"
)
```

文件：

```
asp/asp_parser.rs
```

结构：

```
pub enum AspNode {
    Html(String),
    Script(String),
}
```

执行流程：

```
ASP Page
   │
   ▼
ASP Parser
   │
   ▼
HTML + Script Blocks
   │
   ▼
VBScript Engine
```

------

# 10. Session System

Session 通过 **Cookie + 文件存储** 实现。

SessionID：

```
ASPSESSIONID=abc123
```

存储目录：

```
sessions/
    abc123.json
```

Session 文件：

```
{
  "user": "Tom",
  "login": true
}
```

Session 对象：

```
pub struct Session {
    id: String,
    data: HashMap<String, Value>
}
```

------

# 11. Execution Flow

完整执行流程：

```
HTTP Request
     │
     ▼
Load ASP File
     │
     ▼
ASP Parser
     │
     ├── HTML → Response.Write
     │
     └── Script
             │
             ▼
           Lexer
             │
             ▼
           Parser
             │
             ▼
            AST
             │
             ▼
        Interpreter
             │
             ▼
       Builtins / Objects
             │
             ▼
        Response Buffer
             │
             ▼
        HTTP Response
```

------

# 12. Code Size Strategy

为了避免单文件过大：

| 模块        | 最大行数   |
| ----------- | ---------- |
| lexer       | < 300      |
| parser      | < 400      |
| interpreter | < 400      |
| builtins    | 每个 < 200 |
| objects     | 每个 < 200 |

原则：

```
一个模块一个职责
一个文件不超过 500 行
```

------

# 13. Future Extensions

未来可以扩展：

### COM Object

```
Set obj = Server.CreateObject("ADODB.Connection")
```

### Include

```
<!--#include file="header.asp"-->
```

### Error Handling

```
On Error Resume Next
```

### Class

```
Class User
End Class
```

------

# 14. Summary

核心模块：

```
Lexer
Parser
AST
Interpreter
Runtime
Builtins
ASP Parser
```

执行核心：

```
ASP Page
   ↓
VBScript Engine
   ↓
Runtime Objects
   ↓
HTTP Response
```

该架构：

-   模块清晰
-   易维护
-   易扩展
-   接近真实 ASP 引擎结构