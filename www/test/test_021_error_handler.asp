<%
On Error Resume Next
Response.Write "开始执行<br>"
Result = 1 / 0  ' 故意制造错误
If Err.Number <> 0 Then
    Response.Write "捕获错误: " & Err.Description
    Err.Clear
End If
Response.Write "<br>继续执行"
%>
