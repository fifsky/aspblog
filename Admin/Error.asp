<%
IF session("admin_name")="" Then
Response.Write "<div align='center'><B>对不起你的操作权限不够！或你停留的时间过长，请<a href='index.asp' target='_parent'>重新登陆</a>！</B>"
Response.Write "</div>"
Response.End
End IF
%>