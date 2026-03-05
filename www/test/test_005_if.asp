<%
Dim score
score = 85

If score >= 90 Then
    Response.Write "优秀"
ElseIf score >= 80 Then
    Response.Write "良好"
ElseIf score >= 60 Then
    Response.Write "及格"
Else
    Response.Write "不及格"
End If
%>
