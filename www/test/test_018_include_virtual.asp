<%
' 注意: virtual 需要从网站根目录开始
' 这里使用 file 作为替代演示
%>
<!--#include file="header.asp"-->
<p>Virtual Include 测试内容区域</p>
<%
Response.Write "Virtual 测试: " & "成功"
%>
<!--#include file="footer.asp"-->
