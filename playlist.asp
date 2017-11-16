<!--#include file="conn.asp"-->
<%
songid=FormatSQL(SafeRequest("songid",1))
if songid="" then
response.Write "´íÎó£¡"
response.end
end if
set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from HN_down where id="&songid
rs.open sql,conn,1,3
if songid=""&rs("id")&"" then
response.Write ""&rs("come")&""
response.end
end if
%>