<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim username
    username = Request.Form("username")
    Response.Write "提交的用户名: " & username
Else
%>
<form method="post" action="">
    <input type="text" name="username">
    <input type="submit" value="提交">
</form>
<% End If %>
