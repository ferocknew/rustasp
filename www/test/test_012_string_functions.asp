<%
Dim text
text = "Hello World"

Response.Write "长度: " & Len(text) & "<br>"
Response.Write "左: " & Left(text, 5) & "<br>"
Response.Write "右: " & Right(text, 5) & "<br>"
Response.Write "大写: " & UCase(text) & "<br>"
Response.Write "小写: " & LCase(text)
%>
