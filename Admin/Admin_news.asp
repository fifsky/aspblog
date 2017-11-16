<!--#include file="conn.asp"-->
<!--#include file="Error.asp"-->
<!--#include file="page.asp"-->
<link href=images/admin.css rel=stylesheet>
<META http-equiv=Content-Type content=text/html;charset=GB2312>
<body Style="background-color:#9EB6D8" text="#000000" leftmargin="5" topmargin="5">
<%
work=Request("work")
dome=Request("dome")
types=Request("types")
If Request("IN_yn")="yes" then
  IN_yn=true
 else
  IN_yn=false
 end if
IF work="add" Then
 sql="insert into HN_info (IN_dome,IN_type,IN_yn) values ('"&dome&"','"&types&"',"&IN_yn&")"
 Conn.execute(sql)
 Response.Write "<SCRIPT LANGUAGE='JavaScript'>"
 Response.Write "alert('信息添加成功');"
 Response.Write "</SCRIPT>"
End IF
IF work="del" Then
 sql="delete from HN_info where IN_type="&Request("types")
 Conn.execute(sql)
 Response.Write "<SCRIPT LANGUAGE='JavaScript'>"
 Response.Write "alert('信息删除成功');"
 Response.Write "</SCRIPT>"
End IF
IF work="upedit" Then
 If Request("IN_yn")="yes" then
  IN_yn=true
 else
  IN_yn=false
 end if
 sql="update HN_info set IN_dome='"&Request("dome")&"',IN_yn="&IN_yn&" where IN_type="&Request("types")
 Conn.execute(sql)
 response.redirect Request.ServerVariables("HTTP_REFERER")
End IF
%>
<body>
<%IF Request("work")="up" then
sql="select * from HN_info where IN_type="&Request("types")
set rs=Conn.execute(sql)
%>
<table width="600" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#F7F7F7" class="a2">
<FORM METHOD=POST ACTION="?work=upedit&types=<%=Request("types")%>">
  <tr class="a1">
    <td height="21" colspan="2"><strong>
 <%IF Request("types")=1 Then
  Response.Write "低部广告"
 Elseif Request("types")=2 then
  Response.Write "联系我们"
 elseif Request("types")=3 then
  Response.Write "我的资料"
 elseif Request("types")=0 then
  Response.Write "首页信息"
 End IF%>
    </strong></td>
  </tr>
  <tr class="a3">
    <td width="163" align="right" valign="middle">内容</td>
    <td width="425"> <textarea name="dome" style="display:none"><%=Server.HTMLEncode(rs("IN_dome"))%></textarea>
				  <IFRAME ID="dome" SRC="../ewebedit/ewebeditor.asp?id=dome&style=standard1" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME>
	</td>
  </tr>
  <tr class="a3">
    <td align="right" valign="middle">操作</td>
    <td align="center"><input type="submit" name="Submit" value="提交">
     　　 <input type="reset" name="Submit" value="重置"></td>
  </tr>
  </FORM>
</table>
<%end if%>

</body>