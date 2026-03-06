# Session 功能测试文档

## 测试文件清单

| 编号 | 文件名 | 测试内容 | 状态 |
|------|--------|----------|------|
| 1 | `test_session_001.asp` | 基础读写 | ⬜ |
| 2 | `test_session_002.asp` | SessionID/Timeout | ⬜ |
| 3 | `test_session_003.asp` | Abandon方法 | ⬜ |
| 4 | `test_session_004.asp` | 并发访问 | ⬜ |
| 5 | `test_session_005.asp` | Cookie集成 | ⬜ |

---

## 1. 基础读写测试

**test_session_001.asp**

```asp
<%@ Language=VBScript %>
<html><body>
<h2>Session 基础测试</h2>

<%
' 写入
Session("user") = "Tom"
Session("id") = 123

' 读取
Response.Write "user=" & Session("user") & "<br>"
Response.Write "id=" & Session("id") & "<br>"

' 覆盖
Session("user") = "Jerry"
Response.Write "覆盖后=" & Session("user") & "<br>"
%>

</body></html>
```

**预期结果**: 显示 `user=Tom` → `id=123` → `覆盖后=Jerry`

---

## 2. SessionID/Timeout 测试

**test_session_002.asp**

```asp
<%@ Language=VBScript %>
<html><body>
<h2>Session 属性测试</h2>

<%
' SessionID
Response.Write "SessionID=" & Session.SessionID & "<br>"

' Timeout 默认值
Response.Write "默认Timeout=" & Session.Timeout & "分钟<br>"

' 修改 Timeout
Session.Timeout = 60
Response.Write "修改后Timeout=" & Session.Timeout & "分钟<br>"
%>

</body></html>
```

**预期结果**: SessionID 不为空, 默认 20 分钟, 修改后 60 分钟

---

## 3. Abandon 测试

**test_session_003.asp**

```asp
<%@ Language=VBScript %>
<html><body>
<h2>Session Abandon 测试</h2>

<%
' 设置值
Session("key1") = "value1"
Session("key2") = "value2"

Response.Write "Abandon前:<br>"
Response.Write "key1=" & Session("key1") & "<br>"
Response.Write "key2=" & Session("key2") & "<br>"

' 放弃 Session
Session.Abandon

Response.Write "<br>Abandon后:<br>"
Response.Write "key1=" & Session("key1") & " (应为空)<br>"
Response.Write "key2=" & Session("key2") & " (应为空)<br>"
%>

</body></html>
```

**预期结果**: Abandon 后所有值为空

---

## 4. 并发测试

使用 ab 工具测试:

```bash
ab -n 100 -c 10 http://127.0.0.1:8080/test_session_004.asp
```

---

## 5. Cookie 集成测试

检查浏览器 Cookie:
- Cookie 名: `ASPSESSIONID`
- Cookie 值: SessionID
- 检查是否每次请求都携带

---

## 测试结果汇总表

| 测试项 | 预期结果 | 实际结果 | 状态 |
|--------|----------|----------|------|
| 基础读写 | 正常读写 | - | ⬜ |
| SessionID | 唯一标识 | - | ⬜ |
| Timeout | 可修改 | - | ⬜ |
| Abandon | 清空数据 | - | ⬜ |
| Cookie | 正确设置 | - | ⬜ |
| 并发 | 无冲突 | - | ⬜ |

图例: ⬜ 未测 / 🔄 测试中 / ✅ 通过 / ❌ 失败
