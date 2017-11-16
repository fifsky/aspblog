<%
dim conn,strConn 
set Conn = Server.CreateObject("ADODB.Connection")
strConn = "driver={SQL Server};server=202.101.42.130;uid=sa;pwd=cp041225china050528.;database=caipiaonetcn" 
if err.number<>0 then 
	err.clear
	set conn=nothing
	response.write "数据库连接出错！"
	Response.End
else
	conn.open strConn
	if err then 
		err.clear
		set conn=nothing
		response.write "数据库连接出错！"
		Response.End 
	end if
end if   

sub endConnection()
conn.close
	set conn=nothing
end sub	
%> 