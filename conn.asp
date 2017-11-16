<% 

Function SafeRequest(ParaName,ParaType)
Dim ParaValue
ParaValue=Request(ParaName)
If ParaType=1 then
If not isNumeric(ParaValue) then
Response.write "<center>参数" & ParaName & "必须为数字型，请正确操作！</center>"
Response.end
End if
Else
ParaValue=replace(ParaValue,"'","''")
End if
SafeRequest=ParaValue
End function

Function FormatSQL(strChar)
if IsNull(strChar) Or IsEmpty(strChar) then
FormatSQL=""
else
FormatSQL=replace(strChar,"'","’")
FormatSQL=replace(FormatSQL,"*","×")
FormatSQL=replace(FormatSQL,"?","？")
FormatSQL=replace(FormatSQL,"(","（")
FormatSQL=replace(FormatSQL,")","）")
FormatSQL=replace(FormatSQL,"<","﹤")
FormatSQL=replace(FormatSQL,">","﹥")
FormatSQL=replace(FormatSQL,".","。")
FormatSQL=replace(FormatSQL,";","；")
FormatSQL=replace(FormatSQL,"=","＝")
FormatSQL=replace(FormatSQL,"%","％")
FormatSQL=replace(FormatSQL,"&","※")
end if
End Function 


dim conn   
dim connstr
db="Database/data.mdb"
connstr = "DBQ=" + server.mappath(db) + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"       
set conn=server.createobject("ADODB.CONNECTION")
if err.number<>0 then 
	err.clear
	set conn=nothing
	response.write "数据库连接出错！"
	Response.End
else
	conn.open connstr
	if err then 
		err.clear
		set conn=nothing
		response.write "数据库连接出错！"
		Response.End 
	end if
end if
%>
