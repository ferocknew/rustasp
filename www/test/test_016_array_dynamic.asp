<%
Dim numbers()
ReDim numbers(2)
numbers(0) = 10
numbers(1) = 20
numbers(2) = 30

ReDim Preserve numbers(4)
numbers(3) = 40
numbers(4) = 50

For Each num In numbers
    Response.Write num & "<br>"
Next
%>
