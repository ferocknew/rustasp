<%
' 未定义变量在算术运算中应被当作 0
result = 10 + undefinedVar
Response.Write "10 + undefined = " & result  ' 应显示 10
%>
