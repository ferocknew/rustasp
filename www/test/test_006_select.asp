<%
Dim dayOfWeek
dayOfWeek = 3  ' 假设 3 = 星期三

Select Case dayOfWeek
    Case 1
        Response.Write "星期一"
    Case 2
        Response.Write "星期二"
    Case 3
        Response.Write "星期三"
    Case Else
        Response.Write "其他"
End Select
%>
