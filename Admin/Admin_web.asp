<!--#include file="conn.asp"-->
<!--#include file="Error.asp"-->
<%

Dim Action,FoundErr,ErrMsg
Action=Trim(Request("Action"))
if Action="" then
	call Main()
elseif Action="WebBasicInfoSave" then
	call WebBasicInfoSave()
elseif Action="WebBasicInfoRestore" then
	call WebBasicInfoRestore()
Else
	call Main()
end if
if founderr=true then
	call WriteErrMsg()
End if

%>

<%
Sub Main()
Dim W_ID
Dim W_WebsiteAppe,W_SuppExpl,W_WebSiteName,W_WebSiteUrl,W_StatMastName,W_StatMastEmail,W_StatMastQQ
Dim W_WebSiteKeyword,W_WebSiteIntr,W_WebSiteCopyInfo
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
%>
<link href=images/admin.css rel=stylesheet>
<META http-equiv=Content-Type content=text/html;charset=GB2312>
<body Style="background-color:#9EB6D8" text="#000000" leftmargin="5" topmargin="5">
<table cellpadding="3" cellspacing="1" border="0" width="95%" class="a2">
 <form name="myform" method="post" action="?Action=WebBasicInfoSave&ID=<%=W_ID%>"> <tr class="a1">
    <td colspan="2"><div align="center"><strong>网　站　基　本　设　置</strong></div></td>
  </tr>
   <tr>
     <td colspan="2" class="a3">如果您的网站的设置搞乱了，可以使用<A href="?Action=WebBasicInfoRestore&ID=<%=W_ID%>"><B><font color="#FF0000">还原网站默认设置</font></B></A> </td>
    </tr>
   <tr class="a4">
    <td width="150">网站当前状态：</td>
    <td><input name="W_WebsiteAppe" type="radio" value="True" <%If W_WebsiteAppe="True" Then%>checked<%End If%>>
      打开
        <input type="radio" name="W_WebsiteAppe" value="False" <%If W_WebsiteAppe="False" Then%>checked<%End If%>>
      关闭　<font color="#ff6600">注：维护期间可设置关闭网站</font></td>
  </tr>
  <tr class="a3">
    <td width="150">维护说明：</td>
    <td><textarea name="W_SuppExpl" cols="60" rows="5" class="input_text" id="W_SuppExpl"><%=W_SuppExpl%></textarea>
      <br>
      <font color="#ff6600">注：在网站关闭情况下显示，支持html语法</font></td>
  </tr>
  <tr class="a4">
    <td width="150">网站名称：</td>
    <td><input name="W_WebSiteName" type="text" class="input_text" id="W_WebSiteName" size="40" value="<%=W_WebSiteName%>"></td>
  </tr>
  <tr class="a3">
    <td width="150">网站网址：</td>
    <td><input name="W_WebSiteUrl" type="text" class="input_text" id="W_WebSiteUrl" size="40" value="<%=W_WebSiteUrl%>"></td>
  </tr>
  <tr class="a4">
    <td width="150">站长姓名：</td>
    <td><input name="W_StatMastName" type="text" class="input_text" id="W_StatMastName" size="40" value="<%=W_StatMastName%>"></td>
  </tr>
  <tr class="a3">
    <td width="150">站长邮箱：</td>
    <td><input name="W_StatMastEmail" type="text" class="input_text" id="W_StatMastEmail" size="40" value="<%=W_StatMastEmail%>"></td>
  </tr>
  <tr class="a4">
    <td width="150">站长ＱＱ：</td>
    <td><input name="W_StatMastQQ" type="text" class="input_text" id="W_StatMastQQ" size="40" value="<%=W_StatMastQQ%>"></td>
  </tr>
  <tr class="a3">
    <td width="150">网站关键字：</td>
    <td><textarea name="W_WebSiteKeyword" cols="60" rows="5" class="input_text" id="W_WebSiteKeyword"><%=W_WebSiteKeyword%></textarea></td>
  </tr>
  <tr class="a4">
    <td width="150">网站介绍：</td>
    <td><textarea name="W_WebSiteIntr" cols="60" rows="10" class="input_text" id="W_WebSiteIntr"><%=W_WebSiteIntr%></textarea></td>
  </tr>
  <tr class="a3">
    <td width="150">网站版权信息:</td>
    <td><textarea name="W_WebSiteCopyInfo" cols="60" rows="5" class="input_text" id="W_WebSiteCopyInfo"><%=W_WebSiteCopyInfo%></textarea></td>
  </tr>
  <tr class="a4">
    <td height="40" colspan="2"><div align="center">
      <input type="submit" name="Submit" value="　提　交　" class="input_submit">
    </div></td>
    </tr>
 </form>
</table>
<%End Sub%>


<%
Sub WebBasicInfoSave()
Dim W_ID
Dim W_WebsiteAppe,W_SuppExpl,W_WebSiteName,W_WebSiteUrl,W_StatMastName,W_StatMastEmail,W_StatMastQQ
Dim W_WebSiteKeyword,W_WebSiteIntr,W_WebSiteCopyInfo

W_ID=Trim(Request("ID"))
W_WebsiteAppe=Trim(Request("W_WebsiteAppe"))'网站当前状态
W_SuppExpl=Trim(Request("W_SuppExpl"))'维护说明
W_WebSiteName=Trim(Request("W_WebSiteName"))'网站名称
W_WebSiteUrl=Trim(Request("W_WebSiteUrl"))'网站网址
W_StatMastName=Trim(Request("W_StatMastName"))'站长名称
W_StatMastEmail=Trim(Request("W_StatMastEmail"))'站长Email
W_StatMastQQ=Trim(Request("W_StatMastQQ"))'站长　QQ
W_WebSiteKeyword=Trim(Request("W_WebSiteKeyword"))'网站关键字
W_WebSiteIntr=Trim(Request("W_WebSiteIntr"))'网站介绍
W_WebSiteCopyInfo=Trim(Request("W_WebSiteCopyInfo"))'网站版权信息



If W_WebSiteName="" Then
	FounDerr=True
	ErrMsg=ErrMsg&"<br><li>网站名称不能为空！</li>"
End If


set oRs=server.createobject("adodb.recordset")
sSql="Select * From WebBasicInfo Where W_ID="&W_ID
oRs.Open sSql,Conn,1,3
oRs("W_WebsiteAppe")=W_WebsiteAppe
oRs("W_SuppExpl")=W_SuppExpl
oRs("W_WebSiteName")=W_WebSiteName
oRs("W_WebSiteUrl")=W_WebSiteUrl
oRs("W_StatMastName")=W_StatMastName
oRs("W_StatMastEmail")=W_StatMastEmail
oRs("W_StatMastQQ")=W_StatMastQQ
oRs("W_WebSiteKeyword")=W_WebSiteKeyword
oRs("W_WebSiteIntr")=W_WebSiteIntr
oRs("W_WebSiteCopyInfo")=W_WebSiteCopyInfo

oRs.Update
oRs.Close

Response.Write "<p align=center>修改成功！<script>window.setTimeout(""location.href='?'"",1000);</script></p>"

End Sub
%>


<%
Sub WebBasicInfoRestore()
W_ID=Trim(Request("ID"))
set oRs=server.createobject("adodb.recordset")

sSql="Select * From WebBasicInfo Where W_ID="&W_ID
oRs.Open sSql,Conn,1,3
oRs("W_WebsiteAppe")="True"
oRs("W_SuppExpl")="网站维护中"
oRs("W_WebSiteName")="威涛设计"
oRs("W_WebSiteUrl")="http://www.VTsj.cn"
oRs("W_StatMastName")="威涛"
oRs("W_StatMastEmail")="68747007@qq.com"
oRs("W_StatMastQQ")="68747007"
oRs("W_WebSiteKeyword")="威涛设计"
oRs("W_WebSiteIntr")="欢迎光临威涛设计"
oRs("W_WebSiteCopyInfo")="COPYRIGHT 2006-2008 &copy; VTsj.CN QQ:68747007"



oRs.Update
oRs.Close


Response.Write "<p align=center>还原网站默认设置成功！<script>window.setTimeout(""location.href='?'"",1000);</script></p>"


End Sub
%>