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
    <td colspan="2"><div align="center"><strong>����վ�����������衡��</strong></div></td>
  </tr>
   <tr>
     <td colspan="2" class="a3">���������վ�����ø����ˣ�����ʹ��<A href="?Action=WebBasicInfoRestore&ID=<%=W_ID%>"><B><font color="#FF0000">��ԭ��վĬ������</font></B></A> </td>
    </tr>
   <tr class="a4">
    <td width="150">��վ��ǰ״̬��</td>
    <td><input name="W_WebsiteAppe" type="radio" value="True" <%If W_WebsiteAppe="True" Then%>checked<%End If%>>
      ��
        <input type="radio" name="W_WebsiteAppe" value="False" <%If W_WebsiteAppe="False" Then%>checked<%End If%>>
      �رա�<font color="#ff6600">ע��ά���ڼ�����ùر���վ</font></td>
  </tr>
  <tr class="a3">
    <td width="150">ά��˵����</td>
    <td><textarea name="W_SuppExpl" cols="60" rows="5" class="input_text" id="W_SuppExpl"><%=W_SuppExpl%></textarea>
      <br>
      <font color="#ff6600">ע������վ�ر��������ʾ��֧��html�﷨</font></td>
  </tr>
  <tr class="a4">
    <td width="150">��վ���ƣ�</td>
    <td><input name="W_WebSiteName" type="text" class="input_text" id="W_WebSiteName" size="40" value="<%=W_WebSiteName%>"></td>
  </tr>
  <tr class="a3">
    <td width="150">��վ��ַ��</td>
    <td><input name="W_WebSiteUrl" type="text" class="input_text" id="W_WebSiteUrl" size="40" value="<%=W_WebSiteUrl%>"></td>
  </tr>
  <tr class="a4">
    <td width="150">վ��������</td>
    <td><input name="W_StatMastName" type="text" class="input_text" id="W_StatMastName" size="40" value="<%=W_StatMastName%>"></td>
  </tr>
  <tr class="a3">
    <td width="150">վ�����䣺</td>
    <td><input name="W_StatMastEmail" type="text" class="input_text" id="W_StatMastEmail" size="40" value="<%=W_StatMastEmail%>"></td>
  </tr>
  <tr class="a4">
    <td width="150">վ���ѣѣ�</td>
    <td><input name="W_StatMastQQ" type="text" class="input_text" id="W_StatMastQQ" size="40" value="<%=W_StatMastQQ%>"></td>
  </tr>
  <tr class="a3">
    <td width="150">��վ�ؼ��֣�</td>
    <td><textarea name="W_WebSiteKeyword" cols="60" rows="5" class="input_text" id="W_WebSiteKeyword"><%=W_WebSiteKeyword%></textarea></td>
  </tr>
  <tr class="a4">
    <td width="150">��վ���ܣ�</td>
    <td><textarea name="W_WebSiteIntr" cols="60" rows="10" class="input_text" id="W_WebSiteIntr"><%=W_WebSiteIntr%></textarea></td>
  </tr>
  <tr class="a3">
    <td width="150">��վ��Ȩ��Ϣ:</td>
    <td><textarea name="W_WebSiteCopyInfo" cols="60" rows="5" class="input_text" id="W_WebSiteCopyInfo"><%=W_WebSiteCopyInfo%></textarea></td>
  </tr>
  <tr class="a4">
    <td height="40" colspan="2"><div align="center">
      <input type="submit" name="Submit" value="���ᡡ����" class="input_submit">
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
W_WebsiteAppe=Trim(Request("W_WebsiteAppe"))'��վ��ǰ״̬
W_SuppExpl=Trim(Request("W_SuppExpl"))'ά��˵��
W_WebSiteName=Trim(Request("W_WebSiteName"))'��վ����
W_WebSiteUrl=Trim(Request("W_WebSiteUrl"))'��վ��ַ
W_StatMastName=Trim(Request("W_StatMastName"))'վ������
W_StatMastEmail=Trim(Request("W_StatMastEmail"))'վ��Email
W_StatMastQQ=Trim(Request("W_StatMastQQ"))'վ����QQ
W_WebSiteKeyword=Trim(Request("W_WebSiteKeyword"))'��վ�ؼ���
W_WebSiteIntr=Trim(Request("W_WebSiteIntr"))'��վ����
W_WebSiteCopyInfo=Trim(Request("W_WebSiteCopyInfo"))'��վ��Ȩ��Ϣ



If W_WebSiteName="" Then
	FounDerr=True
	ErrMsg=ErrMsg&"<br><li>��վ���Ʋ���Ϊ�գ�</li>"
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

Response.Write "<p align=center>�޸ĳɹ���<script>window.setTimeout(""location.href='?'"",1000);</script></p>"

End Sub
%>


<%
Sub WebBasicInfoRestore()
W_ID=Trim(Request("ID"))
set oRs=server.createobject("adodb.recordset")

sSql="Select * From WebBasicInfo Where W_ID="&W_ID
oRs.Open sSql,Conn,1,3
oRs("W_WebsiteAppe")="True"
oRs("W_SuppExpl")="��վά����"
oRs("W_WebSiteName")="�������"
oRs("W_WebSiteUrl")="http://www.VTsj.cn"
oRs("W_StatMastName")="����"
oRs("W_StatMastEmail")="68747007@qq.com"
oRs("W_StatMastQQ")="68747007"
oRs("W_WebSiteKeyword")="�������"
oRs("W_WebSiteIntr")="��ӭ�����������"
oRs("W_WebSiteCopyInfo")="COPYRIGHT 2006-2008 &copy; VTsj.CN QQ:68747007"



oRs.Update
oRs.Close


Response.Write "<p align=center>��ԭ��վĬ�����óɹ���<script>window.setTimeout(""location.href='?'"",1000);</script></p>"


End Sub
%>