<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="inc/char.inc"-->
<!--#include file="include.asp"-->
<%
if request("action")="" then response.redirect"gbook.asp?action=show"
if request("action")="post" then
dim bookname
dim bookwhere
dim bookpic
dim bookface
dim bookcontent
dim booktime
dim bookip

dim FoundErr,ErrMsg


bookname=request("name")
bookwhere=request("where")
bookpic=request("pic")
bookface=request("face")
bookshow=request("show")
bookcontent=request("content")
bookip=request.ServerVariables("REMOTE_ADDR")


if bookname="" then

 Response.Write "<SCRIPT LANGUAGE='JavaScript'>"
 Response.Write "alert('�����ǳƲ���Ϊ�գ�');"
 Response.Write "history.go(-1);"
 Response.Write "</SCRIPT>"
	Response.End
end if
if bookcontent="" then

 Response.Write "<SCRIPT LANGUAGE='JavaScript'>"
 Response.Write "alert('�������ݲ���Ϊ�գ�');"
 Response.Write "history.go(-1);"
 Response.Write "</SCRIPT>"
	Response.End
end if

strArr=split(W_BookWorryNeed,"|")  

set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from book"
rs.open sql,conn,1,3
rs.addnew
rs("name")=htmlencode(bookname)
rs("where")=htmlencode(bookwhere)
rs("pic")=bookpic
rs("face")="images/face/"&bookface&".gif"
rs("show")=bookshow
rs("content")=htmlencode(bookcontent)
rs("ip")=bookip
rs("time")=now()
rs.update
rs.close
 Response.Write "<SCRIPT LANGUAGE='JavaScript'>"
 Response.Write "alert('��ϲ�������Ѿ��ɹ��ύ����Ϣ��');"
 Response.Write "window.location.href='gbook.asp?i=show'"
 Response.Write "</SCRIPT>"
Response.end
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���� - <%=WebName%></title>
<meta name="keywords" content="<%=W_WebSiteKeyword%>">
<LINK href="css.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
.STYLE5 {color: #FFFFFF}
-->
</style>
</head>

<body>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div id="Layer1" style="position:absolute; left:345px; top:0px; width:305px; height:84px; z-index:1">
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="333" height="80">
    <param name="movie" value="images/VTsj.swf">
    <param name="quality" value="high">
    <param name="wmode" value="transparent">
    <param name="menu" value="false">
    <embed src="images/VTsj.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="333" height="80"></embed>
  </object>
</div>
<table width="1004" height="631" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="8">
			<img src="images/VTsj_01.jpg" width="493" height="75" alt=""></td>
		<td colspan="5">
			<img src="images/VTsj_02.jpg" width="511" height="75" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="images/VTsj_03.jpg" width="94" height="54" alt=""></td>
		<td colspan="2">
			<img src="images/index.jpg" alt="" width="162" height="54" border="0" usemap="#Map"></td>
		<td colspan="2">
			<img src="images/photo.gif" alt="" width="159" height="54" border="0" usemap="#Map2"></td>
		<td colspan="4">
			<img src="images/diary.gif" alt="" width="156" height="54" border="0" usemap="#Map3"></td>
		<td>
			<img src="images/media.gif" alt="" width="163" height="54" border="0" usemap="#Map4"></td>
		<td colspan="2">
			<img src="images/gbook.gif" alt="" width="158" height="54" border="0" usemap="#Map5"></td>
		<td>
			<img src="images/VTsj_09.jpg" width="112" height="54" alt=""></td>
	</tr>
	<tr>
		<td colspan="2" rowspan="4">
			<img src="images/VTsj_10.jpg" width="117" height="450" alt=""></td>
		<td colspan="2">
			<img src="images/VTsj_11.jpg" width="262" height="20" alt=""></td>
		<td colspan="2">
			<img src="images/VTsj_12.jpg" width="68" height="20" alt=""></td>
		<td colspan="5">
			<img src="images/VTsj_13.jpg" width="427" height="20" alt=""></td>
		<td colspan="2" rowspan="4">
			<img src="images/VTsj_14.jpg" width="130" height="450" alt=""></td>
	</tr>
	<tr>
	  <td width="262" height="414" colspan="2" rowspan="2" valign="top" background="images/VTsj_15.jpg">
<P>
  <%
set rs=server.createobject("adodb.recordset") 
sql="select * from HN_info where IN_type=3"
rs.open sql,Conn,1,1
%>
  <%=rs("IN_dome")%>
  <%
rs.close
set rs=nothing
%>
      </td>
	  <td colspan="2">
			<img src="images/VTsj_16.jpg" width="68" height="171" alt=""></td>
	  <td width="427" height="414" colspan="5" rowspan="2" valign="top" background="images/VTsj_17.jpg"><SCRIPT LANGUAGE="javascript">
function scroll(n)
{temp=n;
Out1.scrollTop=Out1.scrollTop+temp;
if (temp==0) return;
setTimeout("scroll(temp)",80);
}
      </SCRIPT>
              <TABLE WIDTH="100%">
                <TR>
                  <TD VALIGN="TOP" ><DIV ID=Out1 STYLE="width:421; height:370;overflow: hidden">
                    <table width="525"  border="0" cellpadding="5" cellspacing="0">
                      <tr>
                        <td><a href="Gbook.asp?action=sub"><img src="images/bookb.gif" border="0"></a> <a href="Gbook.asp?action=show"><img src="images/booka.gif" border="0"></a> </td>
                      </tr>
                      <tr>
                        <td width="515"><% If Request("action")="show" Then %>
                            <%
   const MaxPerPage=6   '����ҳ����¼�� const ������������
   dim sql 
   dim rs
   dim totalPut   '�ܼ�¼
   dim CurrentPage  '��ǰҳ��
   dim TotalPages   '��ҳ��
   dim i
   dim id
   
%>
                            <%
set rs=server.createobject("adodb.recordset")
If BookExamine=False Then
sql="select * from book order by id desc"
Else
sql="select * from book where bexamine=true order  by id desc"
End If
rs.open sql,Conn,1,1
%>
                            <%
 if rs.eof and rs.bof then
 %>
                            <table width="522"  border="0" align="center" cellpadding="5" cellspacing="1">
                              <tr>
                                <td align="center">��ʱ��û���κ����ԣ�</td>
                              </tr>
                            </table>
                          <%
else
%>
                            <!-- ��ҳ���ܴ����,�ɶ���ʹ�� -->
                            <% 
 if not rs.eof then
  rs.MoveFirst  'ע��ŵ�ǰ����,�����κ�ҳ�����ڵ�һ����¼��
  end if
  rs.pagesize=MaxPerPage  '����ÿҳ�����ʾ��������¼
  If trim(Request("Page"))<>"" then  '��������ҳ�β�Ϊ��
   If IsNumeric(trim(Request("Page"))) = False Then
		GoError "��ҳ���������벻Ҫ��ͼ�ƻ���ϵͳ��"
	End If
	CurrentPage= CLng(request("Page"))   'clng��ת���ɳ�������������,����ֵ����ǰҳ����
	If CurrentPage> rs.PageCount then  '�����ǰҳ�δ�����ҳ��,�����ҳ�θ�ֵ����ǰҳ����
		CurrentPage = rs.PageCount 
	End If 
Else 
	CurrentPage= 1 'һ������������,����ǰҳ��Ϊ��һҳ
End If 


 totalPut=rs.recordcount '���ܼ�¼��ֵ��TOTALPUT
	if CurrentPage<>1 then '�����ǰҳ�������ڵ�һҳ
		if (currentPage-1)*MaxPerPage<totalPut then  '�����ǰҳ��һ����ÿҳ���ļ�¼��С���ܼ�¼�Ļ�
			rs.move(currentPage-1)*MaxPerPage  '��Ե�ǰ��¼������ƶ�
			dim bookmark  '������ǩ����
			bookmark=rs.bookmark '����ǰ��¼�ı�ǩ���ڱ���BOOKMARK��
		end if 
	end if

	dim n,k 
	if (totalPut mod MaxPerPage)=0 then  '�ܼ�¼����ÿҳ����¼������Ľ��Ϊ��ʱ,��N��������ҳ��,�����ټ�һ.
		n= totalPut \ MaxPerPage
	else  
		n= totalPut \ MaxPerPage + 1  
	end if
%>
                            <!-- ��RS��¼ָ��ָ���һ����¼��Ȼ��ʼ�ж��ƶ���¼ʱ����¼��β�Ƿ�Ϊ�գ������Ϊ�ս����ƶ�ָ�룬���������ݶ���ȡ������ֱ����βΪ��ʱ���˳�ѭ�� -->
                            <%
 id=(totalPut-MaxPerPage*(currentPage-1))+1 
i=0
Do While Not rs.EOF and i<MaxPerPage
id=id-1
%>
                            <table width="89%" border="0" align="center" cellpadding="3" cellspacing="1" class="alpha">
                              <tr>
                                <td width="75" align="center">����<%=id%>λ��</td>
                                <td width="376"><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
                                    <tr>
                                      <td width="57">�����ڣ�</td>
                                      <td align="left"><%=rs("time")%></td>
                                      <td width="26" align="center"><img src="images/friend.gif" alt="���ԣ�<%=rs("where")%>" width="18" height="18" /></td>
                                      <td width="30" align="center"><img src="images/ip.gif" alt="IP��<%=rs("ip")%>" width="18" height="18" /></td>
                                    </tr>
                                </table></td>
                              </tr>
                              <tr bgcolor="#666666">
                                <td height="1" colspan="3" ></td>
                              </tr>
                              <tr>
                                <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td>��</td>
                                    </tr>
                                    <tr>
                                      <td align="center"><img src="<%=rs("pic")%>" width="75" height="75" /></td>
                                    </tr>
                                    <tr>
                                      <td height="35" align="center"><strong><%=rs("name")%></strong></td>
                                    </tr>
                                </table></td>
                                <td valign="top"><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                                    <tr>
                                      <td><%if rs("show")=2 then
response.write"����������Ϊ���Ļ���ֻ�й���Ա�ſɲ鿴!"
%>
                                          <%else%>                                      </td>
                                    </tr>
                                    <tr valign="top">
                                      <td height="60"><%=rs("content")%>
                                          <%end if%>                                      </td>
                                    </tr>
                                  </table>
                                    <%
		if rs("reply")<>"" then%>
                                    <table width="80%" border="0" align="left" cellpadding="0" cellspacing="0">
                                      <tr>
                                        <td align="left"><span class="style3">------�����Ǳ�վվ���ظ�������------</span></td>
                                      </tr>
                                      <tr>
                                        <td class="style6 STYLE"><%=htmlencode(rs("reply"))%></td>
                                      </tr>
                                    </table>
                                  <%end if%>
                                </td>
                              </tr>
                          </table>
                          <% i=i+1
  rs.movenext
  loop
  %>
                            <table align="center" >
                              <tr>
                                <td align="center"> �� <font color="#FF0000"><%=currentpage%>/<%=n%></font> ҳ �� <font color="#FF0000"><%=rs.recordcount%></font> ������
                                  <%k=currentPage
		if k<>1 then
			response.write "[<b>"+"<a href='?action=show&page=1'>��ҳ</a></b>] "
			response.write "[<b>"+"<a href='?action=show&page="&cstr(k-1)&"'>��һҳ</a></b>] "
		else
			Response.Write "[��ҳ] [��һҳ]"
		end if
		if k<>n then
			response.write "[<b>"+"<a href='?action=show&page="&cstr(k+1)&"'>��һҳ</a></b>] "
			response.write "[<b>"+"<a href='?action=show&page="&cstr(n)&"'>βҳ</a></b>] "
		else
			Response.Write "[��һҳ] [βҳ]"
		end if
		
		rs.close  
	%>
                                </td>
                              </tr>
                            </table>
                          <%
end if
%>
                            <% EnD iF %>
                            <% If Request("action")="sub" Then %>
                            <form action="?action=post" method="post" name="form" id="form">
                              <table width="100%" border="0" align="center">
                                <tr>
                                  <td width="79" height="23" align="right">�����ǳƣ�</td>
                                  <td colspan="3" align="left"><input name="name" type="text" size="15" class="input2" />
                                    * ����ʹ�����ģ�</td>
                                  <td width="194" rowspan="2" align="left"><img src="images/face/1.gif" name="iface" width="75" height="75" id="iface" /></td>
                                </tr>
                                <tr>
                                  <td height="26" align="right">�������</td>
                                  <td width="71" align="left"><select name="where" size="1" class="input2" id="select2" style="border: 1px solid #E6E6E6">
                                      <option value="δ֪" selected="selected">����</option>
                                      <option value="����">����</option>
                                      <option value="�Ϻ�">�Ϻ�</option>
                                      <option value="���">���</option>
                                      <option value="����">����</option>
                                      <option value="�ӱ�">�ӱ�</option>
                                      <option value="ɽ��">ɽ��</option>
                                      <option value="���ɹ�">���ɹ�</option>
                                      <option value="����">����</option>
                                      <option value="����">����</option>
                                      <option value="������">������</option>
                                      <option value="����">����</option>
                                      <option value="�㽭">�㽭</option>
                                      <option value="����">����</option>
                                      <option value="����">����</option>
                                      <option value="����">����</option>
                                      <option value="ɽ��">ɽ��</option>
                                      <option value="����">����</option>
                                      <option value="����">����</option>
                                      <option value="����">����</option>
                                      <option value="�㶫">�㶫</option>
                                      <option value="����">����</option>
                                      <option value="����">����</option>
                                      <option value="�Ĵ�">�Ĵ�</option>
                                      <option value="����">����</option>
                                      <option value="����">����</option>
                                      <option value="����">����</option>
                                      <option value="����">����</option>
                                      <option value="����">����</option>
                                      <option value="����">����</option>
                                      <option value="�ຣ">�ຣ</option>
                                      <option value="�½�">�½�</option>
                                      <option value="���">���</option>
                                      <option value="����">����</option>
                                      <option value="̨��">̨��</option>
                                      <option value="����">����</option>
                                    </select>
                                  </td>
                                  <td width="66" align="left">ѡ��ͷ��</td>
                                  <td width="90" align="left"><select name="pic" size="1" class="input2" style="border: 1px solid #E6E6E6" onChange="document.images ['iface'].src=options[selectedIndex].value;">
                                      <option value="images/face/1.gif" selected="selected">NO.  01</option>
                                      <%for i=2 to 21%>
                                      <option value="images/face/<%=i%>.gif">NO.
                                        <%if i<10 then
					response.write "0"&i
					else response.write i
					end if
					%>
                                      </option>
                                      <%next%>
                                  </select></td>
                                </tr>
                                <tr>
                                  <td height="26" align="right">�������ʣ�</td>
                                  <td colspan="4" align="left"><input name="show" type="radio" class="input2" value="1" checked="checked" />
                                    ����
                                    <input name="show" type="radio" class="input2" value="2" />
                                    ���Ļ�</td>
                                </tr>
                                <tr>
                                  <td align="right" valign="middle">�������ݣ�<br />
                                      <br />
                                  </td>
                                  <td colspan="4" align="left"><textarea name="content" cols="40" rows="4" id="textarea3" class="input2"></textarea>
                                  </td>
                                </tr>
                                <tr>
                                  <td colspan="5" valign="middle" align="center"><input name="submit" type="submit" class="input2" id="submit3" value="��������" onClick="javascript:return CheckForm();"/>
                                      <input type="reset" name="Submit2" value="�������" class="input2" />
                                      <input type="button" value="��������" name="cmdExit" onClick=" history.back()" class="input2" />
                                  </td>
                                </tr>
                              </table>
                            </form>
                          <%
EnD iF
%>
                        </td>
                      </tr>
                    </table>
                  </DIV></TD>
                </TR>
                <TR>
                  <TD align="right" VALIGN="TOP" ><IMG SRC="images/up.gif" WIDTH="50" HEIGHT="20" onMouseOver="scroll(-1)" onMouseOut="scroll(0)" onMouseDown="scroll(-3)" BORDER="0" ALT="��������ٶȻ���죡"> <IMG SRC="images/down.gif" onMouseOver="scroll(1)" onMouseOut="scroll(0)"  onmousedown="scroll(3)" BORDER="0" WIDTH="50" HEIGHT="20" ALT="��������ٶȻ���죡"> </TD>
                </TR>
        </TABLE></td>
	</tr>
	<tr>
		<td colspan="2">
			<img src="images/VTsj_18.jpg" width="68" height="243" alt=""></td>
	</tr>
	<tr>
	  <td colspan="2" background="images/VTsj_19.jpg" width="262" height="16">&nbsp;</td>
		<td colspan="2">
			<img src="images/VTsj_20.jpg" width="68" height="16" alt=""></td>
		<td colspan="5">
			<img src="images/VTsj_21.jpg" width="427" height="16" alt=""></td>
	</tr>
	<tr>
		<td colspan="7">
			<img src="images/VTsj_22.jpg" width="477" height="51" alt=""></td>
		<td colspan="6" background="images/VTsj_23.jpg"><span class="STYLE5"><%=W_WebSiteCopyInfo%></span></td>
	</tr>
	<tr>
		<td>
			<img src="images/fgf.gif" width="94" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="23" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="139" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="123" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="36" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="32" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="30" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="16" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="78" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="163" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="140" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="18" height="1" alt=""></td>
		<td>
			<img src="images/fgf.gif" width="112" height="1" alt=""></td>
	</tr>
</table>
<map name="Map">
  <area shape="rect" coords="12,9,158,41" href="index.asp">
</map>
<map name="Map2">
  <area shape="rect" coords="9,9,154,41" href="photo.asp">
</map>
<map name="Map3">
  <area shape="rect" coords="8,10,153,41" href="diary.asp">
</map>
<map name="Map4">
  <area shape="rect" coords="10,10,156,41" href="media.asp">
</map>
<map name="Map5">
  <area shape="rect" coords="7,10,151,41" href="gbook.asp">
</map>
</body>
</html>