# jonahfu/vbscript

Rust 实现的 Classic ASP（VBScript 子集）运行时，摆脱 IIS，支持容器化部署。

## 快速开始

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

## Docker Compose 部署

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
      - ./www:/app/www
      - ./runtime/sessions:/app/runtime/sessions
    environment:
      # 基础配置
      - DIRECTORY_LISTING=false
      - HOME_DIR=/app/www
      - INDEX_FILE=index.asp,index.html
      - INDEX_FILE_ENABLE=true
      - PORT=8080
      - ALLOW_PARENT_PATH=false
      - ASP_EXT=asp,asa

      # 调试配置
      - DEBUG=false
      - DETAILED_ERROR=false

      # Session 配置
      - SESSION_STORAGE=memory
      - SESSION_TIMEOUT=20
      - SESSION_DIR=/app/runtime/sessions

      # 日期格式
      - NOW_FORMAT=yyyy/mm/dd hh:nn:ss
      - DATE_FORMAT=yyyy/mm/dd
      - TIME_FORMAT=hh:nn:ss

      # CreateObject 配置
      - CREATE_OBJECT_ENABLE=true
      - CREATE_OBJECT_WHITELIST=Scripting.Dictionary,Scripting.FileSystemObject,MSXML2.XMLHTTP
```

## 环境变量说明

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `HOME_DIR` | Web 根目录 | `/app/www` |
| `PORT` | 服务端口 | `8080` |
| `DIRECTORY_LISTING` | 是否显示目录列表 | `false` |
| `INDEX_FILE` | 默认索引文件 | `index.asp,index.html` |
| `INDEX_FILE_ENABLE` | 索引文件功能开关 | `true` |
| `ALLOW_PARENT_PATH` | 是否支持父路径访问 | `false` |
| `ASP_EXT` | ASP 文件扩展名 | `asp,asa` |
| `DEBUG` | 是否启用调试模式 | `false` |
| `DETAILED_ERROR` | 是否显示详细错误 | `false` |
| `ERROR_PAGE` | 自定义错误页面 | - |
| `SESSION_STORAGE` | Session 存储模式 | `memory` |
| `SESSION_TIMEOUT` | Session 超时(分钟) | `20` |
| `SESSION_DIR` | Session 存储目录 | `/app/runtime/sessions` |
| `NOW_FORMAT` | 日期时间格式 | `yyyy/mm/dd hh:nn:ss` |
| `DATE_FORMAT` | 日期格式 | `yyyy/mm/dd` |
| `TIME_FORMAT` | 时间格式 | `hh:nn:ss` |
| `CREATE_OBJECT_ENABLE` | CreateObject 开关 | `true` |
| `CREATE_OBJECT_WHITELIST` | CreateObject 白名单 | 见配置 |

## 支持的 VBScript 特性

- ✅ 变量声明：Dim, Const, ReDim
- ✅ 流程控制：If/Then/Else, Select Case, For/Next, While/Wend, Do/Loop
- ✅ 过程：Function, Sub
- ✅ 类：Class, Property Get/Let/Set
- ✅ 内置对象：Response, Request, Server, Session
- ✅ COM 对象：Scripting.Dictionary, Scripting.FileSystemObject, MSXML2.XMLHTTP
- ✅ Include 指令

## 不支持的特性

- ❌ COM/ActiveX 组件
- ❌ ADODB 数据库访问
- ❌ Windows-only DLL

## 更多信息

- GitHub: https://github.com/ferocknew/rustasp
- 问题反馈: https://github.com/ferocknew/rustasp/issues
