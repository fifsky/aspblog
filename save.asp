<!--#include file="conn.asp"-->
<% 
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "Cache-Control", "no-cache, must-revalidate"

Function GetRndFileName()
	Dim tmpstr
	randomize
	tmpstr=Int(1000*rnd)
	tmpstr=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&tmpstr
	GetRndFileName=tmpstr
End Function

path="pictemp\"&getrndfilename()&".jpg" 

TotalBytes = Request.TotalBytes  
Set bSourceData = server.createobject("ADODB.Stream")
bSourceData.Open
bSourceData.Type = 1
biData = Request.BinaryRead(TotalBytes)
bSourceData.Write biData 
bSourceData.SaveToFile (server.mappath(path))
set bsourcedata=nothing

set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from HN_pro"
rs.open sql,conn,1,3
rs.addnew
rs("title")=""&getrndfilename()&""
rs("come")=path
rs("class")="2"
rs("time")=date
rs.update
rs.close
set rs=nothing
response.write("fileurl="&replace(path,"\","\\"))



%>

