<!--#include file="conn.asp"-->
<!--#include file="Error.asp"-->
<link href=images/admin.css rel=stylesheet>
<META http-equiv=Content-Type content=text/html;charset=GB2312>
<body Style="background-color:#9EB6D8" text="#000000" leftmargin="5" topmargin="5">
<script language=JavaScript>
<!--
function DoEmpty(params)
{
if (confirm("���Ҫɾ��������¼��ɾ����˼�¼����������ݶ�����ɾ�������޷��ָ���"))
window.location = params ;
}
//-->
</script>
<%
work=Request("work")
Select Case work
 Case "addnews"		'������
  call addnews()

 Case "palynews"	'��ʾ����	
  call palynews()

 Case "editnews"	'�༭����	
  call editnews()

 Case "delnews"		'ɾ������
	sql="delete from HN_pro where id="&Trim(Request("id"))
	set rs=server.createobject("adodb.recordset") 
	rs.open sql,conn,3,3
	response.redirect "?cla="&Request("cla")

 Case "addcla"		'�����������
  call addcla()
 
 Case "palycla"		'�����������
  call palycla()
 
 Case "editcla"		'�������
  Call editcla()

 Case "delcla"		'ɾ�����
	sql="delete from HN_procla where id="&Trim(Request("id"))
	set rs=server.createobject("adodb.recordset") 
	rs.open sql,conn,3,3
	response.redirect "?work=palycla"
 Case Else
  call palynews()
end Select
%>
<%
FUNCTION txtToURL(tekst)
Tekst_temp = tekst
Tekst_temp = Replace(Tekst_temp,"<","&lt;")
Tekst_temp = Replace(Tekst_temp,">","&gt;")
Tekst_temp = Replace(Tekst_temp,"chr(13)","<BR>")
Tekst_temp = Replace(Tekst_temp,"#","<BR>")

Tekst_temp = Replace(Tekst_temp,Chr(13),"<BR>")
Tekst_temp = Replace(Tekst_temp,"[","<B>")
Tekst_temp = Replace(Tekst_temp,"]","</B>")
txtToURL = Tekst_temp
END FUNCTION
%>
<%
'------------------------------
'��������ģ��
'------------------------------
sub addnews()
IF Request("add")<>"yes" Then
%>
<table width="650" height="363" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#e4e4e4" class="a2">
    <tr>
      <td height="24" align="center" valign="middle" bgcolor="#FFFFFF" class="a1">ͼƬ<span class="style1">¼��</span></td>
    </tr>
    <tr>
      <td width="100%" height="142">
        <form name="add" method="POST" action="?work=addnews" >
          <div align="center">          
            <table width="100%" height="266" border="0" cellpadding="3" cellspacing="1" bgcolor="#e4e4e4" class="a2">
              <tr class="a3">
                <td width="11%" height="23" align="center">ͼƬ����</td>
                <td height="23"><input type="text" name="title" size="50" >
      * &nbsp; </td>
              </tr>
              <input type="hidden" name="auth" size="50">
              <tr class="a3">
                <td width="11%" height="20" align="center">ͼƬ��ַ</td>
                <td height="20"><input type="text" name="come" size="50">
      *<br>
      <iframe border=0 frameborder=0 framespacing=0 height=22 width=300 marginheight=0 marginwidth=0 name=new_date noresize scrolling=no  src="upload.htm" vspale="0"></iframe></td>
              </tr>
              <tr class="a3">
                <td width="11%" height="25" align="center">ͼƬ���</td>
                <td height="25"><%
cla=cint(replace(request("cla"),"'",""))
%>
                    <select name="clas" id="clas">
                      <%set rs1=server.createobject("adodb.recordset")
			sql="select * from HN_procla where id="&cla&""
			rs1.open sql,conn,1,1
			do while Not rs1.eof
			%>
                      <option value="<%=rs1("id")%>">
                      <%if rs1("class")<>"" then%>
                      <%=rs1("class")%>
                      <%else%>
                      <%=rs1("MidCls")%>
                      <%end if%>
                      </option>
                      <%
			rs1.movenext
			loop
			rs1.close
			set rs1=nothing
            %>
                    </select></td>
              </tr>
              <tr class="a3">
                <td width="11%" height="90" align="center">ͼƬ˵��</td>
                <td height="90"><textarea name="demo" style="display:none"></textarea>
                    <iframe id="demo" src="../ewebedit/ewebeditor.asp?id=demo&style=standard1" frameborder="0" scrolling="no" width="550" height="350"></iframe>
                </td>
              </tr>
              <tr class="a3">
                <input type="hidden" name="add" value="yes">
                <td height="29" colspan="2" align="center"><input type="submit" value="����ͼƬ" name="submit">
&nbsp;
      <input type="button" name="Submit" onClick=history.go(-1) value="����">
                </td>
              </tr>
            </table>
          </div>
        </form>
      </td>
    </tr>
</table>
<%else
dim title,name,email,picUrl,demo,kinds,come,auth,sql,clas
title=request("title")
picUrl=request("picUrl")
For i = 1 To Request.Form("demo").Count
	sContent = sContent & Request.Form("demo")(i)
Next
demo=sContent
clas=request("clas")
come=request("come")
kinds=request("kinds")
if len(demo)=0 then 
	response.write"<center><font color=#ff0000 size=2>���ݲ���Ϊ�գ�</font></center>"
	response.write"<center><font size=2><a href=# onclick=history.go(-1)>�����ﷵ��</a></font></center>"
	response.end
end if
if len(picUrl)=0 then picUrl=" "
if len(come)=0 then come=" "

set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from HN_pro"
rs.open sql,conn,1,3
rs.addnew
rs("title")=title
rs("demo")=demo
rs("come")=come
rs("class")=clas
rs("time")=date
rs.update
rs.close
set rs=nothing
Response.Write "<br>���ӳɹ���<meta http-equiv=refresh content=""0;URL=?work=palynews&cla="&clas&""">"
end if
end sub
%>
<%
sub palynews()
%>
<%
cla=cint(replace(request("cla"),"'",""))
%>
<table width="650" border="0" align="center" cellpadding="3" cellspacing="1" class="a2">
  <tr class="a1">
    <td width="571" height="20" align="center" valign="middle"><span class="tetle_top2">ͼƬ�б�</span></td>
    <td width="79" align="center" valign="middle"><input type="button" name="Submit" value="����ͼƬ" onClick="window.location='?work=addnews&cla=<%= cla %>';">
    </a></td>
  </tr>
  <tr>
    <td colspan="2">
<table border=0 cellspacing=0 cellpadding=0 width=100% class=main_tab>
<%
cla=cint(replace(request("cla"),"'",""))
%>
<tbody> 
        <tr> 
          <td align=center valign=top>
<TABLE width="98%" align="center" cellpadding="3" cellspacing="1" border="0" class="a2" bgcolor="#E7E7E7">
<TR align="center" valign="middle">
	<TD width="5%" height="22" class="a1" >ID</TD>
	<TD width="57%" class="a1">����</TD>
	<TD width="18%" class="a1">����</TD>
	<TD colspan="2" class="a1">��ز���</TD>
</TR>
<%
	set rs=server.createobject("adodb.recordset") 
	sql="select * from HN_pro where class="&cla&" order by id desc" 
	rs.open sql,conn,1,1
	if rs.eof then 
	Response.Write "����ͼƬ��û���ҵ����<br><br>"
	else

proCount=rs.recordcount
	rs.PageSize=20
		if not IsEmpty(Request("ToPage")) then
	       ToPage=CInt(Request("ToPage"))
		   if ToPage>rs.PageCount then
					rs.AbsolutePage=rs.PageCount
					intCurPage=rs.PageCount
		   elseif ToPage<=0 then
					rs.AbsolutePage=1
					intCurPage=1
				else
					rs.AbsolutePage=ToPage
					intCurPage=ToPage
				end if
			else
			        rs.AbsolutePage=1
					intCurPage=1
		  end if
	intCurPage=CInt(intCurPage)
	For i = 1 to rs.PageSize
	if rs.EOF then     
	Exit For 
	end if '����for next ѭ�����ζ�����¼
%><TR align="center" valign="middle" class="a3">
	<TD><font color="#99B4CC"><%=rs("id")%></font></TD>
	<TD align="left"><a href="../photo.asp" target="_blank"><span class="unnamed2">

	<%=rs("title")%></span></a></TD>
	<TD align="left">[<%=rs("time")%>]</TD>
	<TD width="10%"><a href="?work=delnews&cla=<%=cla%>&id=<%=rs("id")%>">ɾ��</a></TD>
	<TD width="10%"><a href="?work=editnews&id=<%=rs("id")%>">�༭</a></TD>
</TR>
<%
rs.MoveNext 
next%>
<form name="form1" method="post" action="?cla=<%=cla%>">
  <span class="tebfont1">�ܹ���<font color="#ff0000"><%=rs.PageCount%></font>ҳ, <font color="#ff0000"><%=proCount%></font>����Ʒ, ��ǰҳ��<font color="#ff0000"><%=intCurPage%> </font><%if intCurPage<>1 then
		  %><a href="?cla=<%=cla%>">��ҳ</a>|<a href="?cla=<%=cla%>&ToPage=<%=intCurPage-1%>">��һҳ</a>|<% end if
if intCurPage<>rs.PageCount then %><a href="?cla=<%=cla%>&ToPage=<%=intCurPage+1%>">��һҳ</a>|<a href="?cla=<%=cla%>&ToPage=<%=rs.PageCount%>"> ���ҳ</a>|
         <% end if%>��ת��</span><input name="topage" type="text" size="2" value="<%=intCurPage%>"> <input type="submit" name="go" value="ת��"></form>
<% 
rs.close 
Set rs = Nothing
end if
%></table>
</td>
</tr>
</tbody>
</table>
</td>
  </tr>
</table>
<%
end sub
%>
<%
'------------------------
'�༭����
'------------------------
sub editnews
if request("edit")<>"yes" and Request("id")<>"" then
dim sql
dim id
id=Request("id")
set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from HN_pro where id="&id
rs.open sql,conn,1,1
%>
<table width="650" height="278" border="0" align="center" cellpadding="3" cellspacing="1" class="a2" >
      <tr>
        <td height="29" align="center" valign="middle" class="a1">ͼƬ�༭</td>
      </tr>
    <tr>
      <td width="100%" height="237">
              <table width="100%" height="267" border="0" align="center" cellpadding="3" cellspacing="1" class="a2">
                <form method="POST" name="add" action="?work=editnews&id=<%=id%>">
				<tr class="a3"> 
                  <td width="13%" height="31" align="center">ͼƬ����</td>
                  <td height="31"> 
                    <input type="text" name="title" size="50" value="<%=rs("title")%>">
*&nbsp;</td>
                </tr>

     
                <tr class="a3"> 
                  <td width="13%" height="27" align="center">ͼƬ��ַ</td>
                  <td height="27"> 
                    <input type="text" name="come" size="50" value="<%=rs("come")%>">
                  *</td>
                </tr>
                <tr class="a3"> 
                <td width="13%" height="27" align="center">ͼƬ���</td>
                <td height="27">
				<%
			set rss=server.createobject("adodb.recordset") 
			sql="select * from HN_procla where MidCls='' or Clid=0 order by descs ASC" 
			rss.open sql,conn,1,1
			%>
			<select name="clas">
			<%do while not rss.eof%>
                  <option value="<%=rss("id")%>" <%if rs("class")=rss("id") then %>selected<%end if%>><%=rss("class")%></option>
			<%
			set rs1=server.createobject("adodb.recordset")
			sql="select * from HN_procla where MidCls<>'' and Clid="&rss("id")
			rs1.open sql,conn,1,1
			do while Not rs1.eof
			%>
			<option value="<%=rs1("id")%>" <%if rs("class")=rs1("id") then %>selected<%end if%>>��<%=rs1("MidCls")%></option>
			<%
			rs1.movenext
			loop
			rs1.close
			set rs1=nothing
            %>
			<%
			rss.movenext
			loop
			rss.close
			set rss=nothing%>
			</select></td>
                </tr>
                <tr class="a3"> 
                  <td width="13%" height="90" align="center">ͼƬ˵��</td>
                  <td height="90"> 
                    <textarea name="demo" style="display:none"><%=Server.HTMLEncode(rs("demo")&"")%></textarea>
				  <IFRAME ID="demo" SRC="../ewebedit/ewebeditor.asp?id=demo&style=standard1" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME>                  </td>
                </tr>
                <tr class="a3">
                <td colspan="2" align="center" valign="middle">
                <INPUT TYPE="hidden" name="edit" value="yes">
                <input type="submit" value="ȷ���޸�">&nbsp;<input type="button" name="Submit" onclick=history.go(-1) value="����">
                </td>
                </tr></form>
        </table>
      </td>
    </tr>
</table>
  </center>
</div>
<%rs.close
set rs=nothing
else
id=Request("id")
set rs=server.createobject("adodb.recordset") 
sql="select * from HN_pro where id="&id
rs.open sql,conn,3,3
title=request("title")
auth=request("auth")
come=request("come")
clas=request("clas")
For i = 1 To request("demo").Count
	sContent = sContent & request("demo")(i)
Next
rs("demo")=sContent
rs("title")=title
rs("come")=come
rs("class")=clas '�����������
rs.update
rs.close
set rs=nothing
Response.Write "<br>���Ÿ��³ɹ���������Զ�����<meta http-equiv=refresh content=""1;URL=?work=palynews&cla="&clas&""">"
end if
end sub
%>
<%
'------------------------
'�����������
'------------------------
sub addcla()
cla=Request("cla")
if cla="yes" then
set rs=server.createobject("adodb.recordset") 
sql="select * from HN_procla"
rs.open sql,conn,3,3
rs.addnew
 IF Request("Clid")<>"" Then
  RS("MidCls")=Request("class")
  RS("Clid")=Request("Clid")
 Else
  RS("class")=Request("class")
 End if
 IF Request("MidCls")<>"" Then RS("MidCls")=Request("MidCls")
 IF Request("Descs")<>"" Then RS("Descs")=Request("Descs")
rs.update
Response.Write "<SCRIPT LANGUAGE=JavaScript>alert('���ӳɹ�')</SCRIPT>"
Response.Redirect "?work=palycla"
end if
%>
<TABLE  width="300" border="0" align="center"  cellpadding="3" cellspacing="1" bgcolor="#E4E4E4" class="a2" style="font-family:����;font-size: 9pt;">
<FORM METHOD=POST ACTION="?work=addcla">
<TR class="a1">
  <TD height="22" colspan="2" align="center" valign="middle">���������</TD>
  </TR>
<%
if Request("CLid")<>"" then
%>
<TR>
  <TD align="right" valign="middle" bgcolor="#F7F7F7">�������ࣺ</TD>
  <TD align="left" valign="middle" bgcolor="#FFFFFF">
  <select name="CLid">
<%
set rs=server.createobject("adodb.recordset") 
sql="select * from HN_procla where MidCls='' or Clid=0 order by descs ASC"
rs.open sql,conn,1,1
DO WHILE Not rs.eof
%>
<option value="<%=rs("id")%>" <%IF cint(Request("Clid"))=rs("id") Then%>selected<%End IF%>><%=rs("class")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
  </select></TD>
</TR>
<%End if%>
<TR>

	<TD align="right" valign="middle" bgcolor="#F7F7F7">�������</TD>
	<TD width="209" align="left" valign="middle" bgcolor="#FFFFFF"><INPUT TYPE="text" NAME="class"></TD>
</TR>
<TR>
  <TD bgcolor="#F7F7F7" align="right" valign="middle">����</TD>
  <TD bgcolor="#FFFFFF" align="left" valign="middle"><input name="Descs" type="text" id="order" size="5"></TD>
</TR>
<TR bgcolor="#FFFFFF">
  <INPUT TYPE="hidden" name="cla" value="yes">
	<TD colspan="2" align="center" valign="middle" >
	  <INPUT TYPE="submit" value="�ύ">
	  &nbsp;<input type="button" name="Submit" onclick=history.go(-1) value="����"></TD>
</TR>
</FORM>
</TABLE>
<%
end sub
%>
<%
'-------------------------
'��ʾ�������
'-------------------------
sub palycla()

set rs=server.createobject("adodb.recordset") 
sql="select * from HN_procla where MidCls='' or Clid=0 order by descs ASC" 
rs.open sql,conn,1,1
%>
<table  width="500" border="0" align="center"  cellpadding="0" cellspacing="1" bgcolor="#E4E4E4" class="a2" style="font-family:����;font-size: 9pt;">
  <tr bgcolor="#F7F7F7" class="a1">
    <td height="22" colspan="5" align="center" valign="middle">ͼƬ���� </td>
  </tr>
  <tr>
    <td width="8%" height="18" align="center" valign="middle" bgcolor="#F7F7F7">ID</td>
    <td width="57%" align="center" valign="middle" bgcolor="#F7F7F7">�����</td>
    <td bgcolor="#F7F7F7" align="center" valign="middle" colspan="3">��ز���</td>
  </tr>
  <%if rs.eof then%>
  <tr bgcolor="#FFFFFF">
    <td height="17" colspan="5" align="center" valign="middle">��ʱû�����</td>
  </tr>
  <%else%>
  <%do while not rs.eof
%>
  <tr bgcolor="#FFFFFF">
    <td height="20" align="center" valign="middle"><%=rs("id")%></td>
    <td align="left" valign="middle"><b><%=rs("class")%></b></td>
    <td width="11%" align="center" valign="middle"><a href="javascript:DoEmpty('?work=delcla&id=<%=rs("id")%>');"></a></td>
    <td width="10%" align="center" valign="middle"><a href="?work=editcla&id=<%=rs("id")%>"></a></td>
    <td width="14%" align="center" valign="middle" bgcolor="#FFFFFF"><a href="?work=addcla&CLid=<%=rs("id")%>">����С��</a></td>
  </tr>
  <%
set rs1=server.createobject("adodb.recordset") 
sql="select * from HN_procla where  MidCls<>'' and Clid="&rs("id") 
rs1.open sql,conn,1,1
if Not rs1.eof then
do while Not rs1.eof
%>
  <tr bgcolor="#FFFFFF">
    <td height="20" align="center" valign="middle"><%=rs1("id")%></td>
    <td align="left" valign="middle"> �� <a href="?work=palynews&cla=<%=rs1("id")%>" target="main"><%=rs1("Midcls")%></a> <a href="?work=addnews&cla=<%=rs1("id")%>">[����]</a></td>
    <td width="11%" align="center" valign="middle"><a href="javascript:DoEmpty('?work=delcla&id=<%=rs1("id")%>');">ɾ��</a> </td>
    <td width="10%" align="center" valign="middle"><a href="?work=editcla&id=<%=rs1("id")%>">�༭</a></td>
    <td width="14%" align="center"><a href="?work=palynews&cla=<%=rs1("id")%>" target="main">[�鿴]</a></td>
  </tr>
  <%
rs1.movenext
loop
Else
%>
  <tr align="left">
    <td colspan="5"><a href="?work=addcla&CLid=<%=rs("id")%>">����С��</a></td>
  </tr>
  <%
End if
rs1.close
set rs1=nothing
rs.movenext
loop
end if
%>
</table>
<%rs.close
set rs=nothing
end sub
%>
<%
'---------------------
'�༭���
'---------------------
sub editcla()

if Request("edit")="yes" then
 set rs=server.createobject("adodb.recordset") 
 sql="select * from HN_procla where id="&Request("id")
 rs.open sql,conn,3,3
 IF Request("Clid")<>"" Then
  RS("MidCls")=Request("class")
  RS("Clid")=Request("Clid")
 Else
  RS("class")=Request("class")
 End if
 IF Request("MidCls")<>"" Then RS("MidCls")=Request("MidCls")
 IF Request("Descs")<>"" Then RS("Descs")=Request("Descs")
rs.update

Response.Write "<SCRIPT LANGUAGE=JavaScript>alert('���³ɹ�')</SCRIPT>"
Response.Redirect "?work=palycla"
else

set rs=server.createobject("adodb.recordset") 
sql="select * from HN_procla where id="&Request("id")
rs.open sql,conn,1,1
%>

<TABLE  width="300" border="0" align="center"  cellpadding="3" cellspacing="1" bgcolor="#E4E4E4" class="a2" style="font-family:����;font-size: 9pt;">
<FORM METHOD=POST ACTION="?work=editcla">
<TR align="center" class="a1">
  <TD height="22" colspan="2" valign="middle">�༭���</TD>
</TR>
<%
if rs("CLid")<>"" and rs("MidCls")<>"" then
%>
<TR>
  <TD align="right" valign="middle" bgcolor="#F7F7F7">�������ࣺ</TD>
  <TD align="left" valign="middle" bgcolor="#FFFFFF">
  <select name="CLid">
<%
set rs1=server.createobject("adodb.recordset") 
sql1="select * from HN_procla where MidCls='' or Clid=0 order by descs ASC"
rs1.open sql1,conn,1,1
DO WHILE Not rs1.eof
%>
<option value="<%=rs1("id")%>" <%IF rs("CLid")=rs1("id") Then%>selected<%End IF%>><%=rs1("class")%></option>
<%
rs1.movenext
loop
rs1.close
set rs1=nothing
%>
  </select></TD>
</TR>
<TR>
	<TD width="71" align="right" valign="middle" bgcolor="#F7F7F7">�������</TD>
	<TD width="216" align="left" valign="middle" bgcolor="#FFFFFF"><INPUT TYPE="text" NAME="class" value="<%=rs("MidCls")%>"></TD>
</TR>
<%Else%>
<TR>
	<TD width="71" align="right" valign="middle" bgcolor="#F7F7F7">�������</TD>
	<TD width="216" align="left" valign="middle" bgcolor="#FFFFFF"><INPUT TYPE="text" NAME="class" value="<%=rs("class")%>"></TD>
</TR>
<%End if%>

<TR>
  <TD bgcolor="#F7F7F7" align="right" valign="middle">����</TD>
  <TD bgcolor="#FFFFFF" align="left" valign="middle"><input name="Descs" type="text" id="Descs" value="<%=rs("Descs")%>" size="5"></TD>
</TR>
<TR bgcolor="#FFFFFF">
  <INPUT TYPE="hidden" name="edit" value="yes">
<INPUT TYPE="hidden" name="id" value="<%=rs("id")%>">
	<TD colspan="2" align="center" valign="middle" ><INPUT name="����" TYPE="submit" value="�ύ">
	  &nbsp;<input type="button" name="Submit" onclick=history.go(-1) value="����"></TD>
</TR>
</FORM>
</TABLE>
<%rs.close
set rs=nothing
end if
end sub
%>