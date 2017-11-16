<%
set oRs=server.createobject("adodb.recordset")

sSql="Select Top 1 * From WebBasicInfo"
oRs.Open sSql,Conn,1,1
If Not oRs.Eof Then
	W_ID=oRs("W_ID")
	W_WebsiteAppe=oRs("W_WebsiteAppe")
	W_SuppExpl=oRs("W_SuppExpl")
	W_WebSiteName=oRs("W_WebSiteName")
	W_WebSiteUrl=oRs("W_WebSiteUrl")
	W_StatMastName=oRs("W_StatMastName")
	W_StatMastEmail=oRs("W_StatMastEmail")
	W_StatMastQQ=oRs("W_StatMastQQ")
	W_WebSiteKeyword=oRs("W_WebSiteKeyword")
	W_WebSiteIntr=oRs("W_WebSiteIntr")
	W_WebSiteCopyInfo=oRs("W_WebSiteCopyInfo")

End If
oRs.Close
If W_WebsiteAppe="False" Then
	Response.Write "<table width='100%' height='100%'  border='0' cellpadding='0' cellspacing='0'><tr><td style='font-size:11pt'><div align='center'><font color='#ff0000'>"
	Response.Write W_SuppExpl
	Response.Write "</font></div></td></tr></table>"
	Response.End
End If
Dim WebTitle,WebName
WebName=W_WebSiteName&"-"&W_WebSiteUrl
WebTitle=""

%>

