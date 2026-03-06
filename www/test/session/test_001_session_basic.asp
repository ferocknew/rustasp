<%@ Language="VBScript" %>
<%
' Session 基础功能测试
' 测试 Session 变量的设置、读取、修改
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Session 基础功能测试</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 50px auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .test-container {
            background: white;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #333;
            border-bottom: 2px solid #007bff;
            padding-bottom: 10px;
        }
        .test-section {
            margin: 20px 0;
            padding: 15px;
            border: 1px solid #ddd;
            border-radius: 5px;
            background-color: #f9f9f9;
        }
        .test-section h3 {
            margin-top: 0;
            color: #007bff;
        }
        .result {
            padding: 10px;
            margin: 10px 0;
            border-radius: 4px;
        }
        .success {
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }
        .info {
            background-color: #d1ecf1;
            border: 1px solid #bee5eb;
            color: #0c5460;
        }
        .error {
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0;
        }
        th, td {
            padding: 10px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #007bff;
            color: white;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .action-link {
            display: inline-block;
            margin: 5px;
            padding: 8px 15px;
            background-color: #007bff;
            color: white;
            text-decoration: none;
            border-radius: 4px;
        }
        .action-link:hover {
            background-color: #0056b3;
        }
        .code {
            font-family: 'Courier New', monospace;
            background-color: #f4f4f4;
            padding: 2px 6px;
            border-radius: 3px;
        }
    </style>
</head>
<body>
    <div class="test-container">
        <h1>🧪 Session 基础功能测试</h1>

        <div class="test-section">
            <h3>📋 会话信息</h3>
            <table>
                <tr>
                    <th>属性</th>
                    <th>值</th>
                </tr>
                <tr>
                    <td><strong>SessionID</strong></td>
                    <td><span class="code"><%= Session.SessionID %></span></td>
                </tr>
                <tr>
                    <td><strong>Timeout</strong></td>
                    <td><span class="code"><%= Session.Timeout %> 分钟</span></td>
                </tr>
            </table>
        </div>

        <%
        ' 测试 1: 设置 Session 变量
        If Request.QueryString("action") = "set" Then
            Session("username") = "测试用户"
            Session("login_time") = Now()
            Session("visit_count") = 0
            Session("test_value") = "这是一个测试值"
        %>
        <div class="result success">
            ✅ <strong>测试 1 - 设置 Session 变量</strong><br>
            已设置以下 Session 变量：
            <ul>
                <li><span class="code">Session("username")</span> = "测试用户"</li>
                <li><span class="code">Session("login_time")</span> = <%= Now() %></li>
                <li><span class="code">Session("visit_count")</span> = 0</li>
                <li><span class="code">Session("test_value")</span> = "这是一个测试值"</li>
            </ul>
        </div>

        <%
        ' 测试 2: 读取 Session 变量
        ElseIf Request.QueryString("action") = "read" Then
        %>
        <div class="result info">
            📖 <strong>测试 2 - 读取 Session 变量</strong><br>
            当前 Session 变量值：
            <ul>
                <li><span class="code">Session("username")</span> = "<%= Session("username") %>"</li>
                <li><span class="code">Session("login_time")</span> = <%= Session("login_time") %></li>
                <li><span class="code">Session("visit_count")</span> = <%= Session("visit_count") %></li>
                <li><span class="code">Session("test_value")</span> = "<%= Session("test_value") %>"</li>
            </ul>
        </div>

        <%
        ' 测试 3: 修改 Session 变量
        ElseIf Request.QueryString("action") = "modify" Then
            Session("visit_count") = Session("visit_count") + 1
            Session("last_visit") = Now()
        %>
        <div class="result success">
            ✏️ <strong>测试 3 - 修改 Session 变量</strong><br>
            已修改 Session 变量：
            <ul>
                <li><span class="code">Session("visit_count")</span> 从 <%= Session("visit_count") - 1 %> 增加到 <%= Session("visit_count") %></li>
                <li><span class="code">Session("last_visit")</span> 设置为 <%= Now() %></li>
            </ul>
        </div>

        <%
        ' 测试 4: 删除 Session 变量
        ElseIf Request.QueryString("action") = "remove" Then
            Session.Contents.Remove("test_value")
        %>
        <div class="result success">
            🗑️ <strong>测试 4 - 删除 Session 变量</strong><br>
            已删除 <span class="code">Session("test_value")</span>
        </div>

        <%
        ' 测试 5: 清空所有 Session 变量
        ElseIf Request.QueryString("action") = "clear" Then
            Session.Contents.RemoveAll()
        %>
        <div class="result success">
            🧹 <strong>测试 5 - 清空所有 Session 变量</strong><br>
            已调用 <span class="code">Session.Contents.RemoveAll()</span>
        </div>

        <%
        ' 测试 6: 销毁 Session
        ElseIf Request.QueryString("action") = "abandon" Then
            Session.Abandon()
        %>
        <div class="result error">
            💥 <strong>测试 6 - 销毁 Session</strong><br>
            已调用 <span class="code">Session.Abandon()</span><br>
            下次刷新将获得新的 SessionID
        </div>
        <%
        End If
        %>

        <div class="test-section">
            <h3>🎮 测试操作</h3>
            <p>
                <a href="?action=set" class="action-link">1️⃣ 设置 Session</a>
                <a href="?action=read" class="action-link">2️⃣ 读取 Session</a>
                <a href="?action=modify" class="action-link">3️⃣ 修改 Session</a>
            </p>
            <p>
                <a href="?action=remove" class="action-link">4️⃣ 删除单个变量</a>
                <a href="?action=clear" class="action-link">5️⃣ 清空所有</a>
                <a href="?action=abandon" class="action-link">6️⃣ 销毁 Session</a>
            </p>
            <p>
                <a href="?" class="action-link">🔄 刷新页面</a>
                <a href="../" class="action-link">🔙 返回测试列表</a>
            </p>
        </div>

        <div class="test-section">
            <h3>📊 当前 Session 状态</h3>
            <table>
                <tr>
                    <th>变量名</th>
                    <th>值</th>
                    <th>类型</th>
                </tr>
                <tr>
                    <td><span class="code">username</span></td>
                    <td><%= Session("username") %></td>
                    <td><%= TypeName(Session("username")) %></td>
                </tr>
                <tr>
                    <td><span class="code">login_time</span></td>
                    <td><%= Session("login_time") %></td>
                    <td><%= TypeName(Session("login_time")) %></td>
                </tr>
                <tr>
                    <td><span class="code">visit_count</span></td>
                    <td><%= Session("visit_count") %></td>
                    <td><%= TypeName(Session("visit_count")) %></td>
                </tr>
                <tr>
                    <td><span class="code">last_visit</span></td>
                    <td><%= Session("last_visit") %></td>
                    <td><%= TypeName(Session("last_visit")) %></td>
                </tr>
                <tr>
                    <td><span class="code">test_value</span></td>
                    <td><%= Session("test_value") %></td>
                    <td><%= TypeName(Session("test_value")) %></td>
                </tr>
            </table>
            <p><strong>Session 变量总数:</strong> <%= Session.Contents.Count %></p>
        </div>

        <div class="test-section">
            <h3>📝 测试说明</h3>
            <p>此测试页面验证以下 Session 功能：</p>
            <ul>
                <li>✅ 设置 Session 变量</li>
                <li>✅ 读取 Session 变量</li>
                <li>✅ 修改 Session 变量</li>
                <li>✅ 删除单个 Session 变量</li>
                <li>✅ 清空所有 Session 变量</li>
                <li>✅ 销毁 Session (Abandon)</li>
                <li>✅ 获取 SessionID</li>
                <li>✅ 获取 Timeout</li>
                <li>✅ 获取变量数量</li>
            </ul>
        </div>
    </div>
</body>
</html>
