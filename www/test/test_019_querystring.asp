<%
Dim name
name = Request.QueryString("name")
If name = "" Then name = "访客"
Response.Write "欢迎, " & name
%>

<p><a href="?name=张三">测试: name=张三</a></p>
