<!--#include file="conn.asp"-->
<!-- #include file="../inc/char.inc" -->
<!--#include file="Error.asp"-->
<link href=images/admin.css rel=stylesheet>
<META http-equiv=Content-Type content=text/html;charset=GB2312>
<body Style="background-color:#9EB6D8" text="#000000" leftmargin="5" topmargin="5">
<%
i=Request("i")
if i="" then
call a()
else
	Select Case i
	 Case "1"
	  call a()

	 Case "2"	
	  call b()
	
	 Case "3"
	  call c()
	end Select
end if
%>
<%
'------------------------
'a
'------------------------
sub a()
%>
<%
   const MaxPerPage=5   '����ҳ����¼�� const ������������
   dim sql 
   dim rs
   dim totalPut   '�ܼ�¼
   dim CurrentPage  '��ǰҳ��
   dim TotalPages   '��ҳ��
   dim i
%>

<%
dim filename,id
set rs=server.createobject("adodb.recordset")
sql="select * from book order by id desc"
rs.open sql,conn,1,1
%>
<script language="JavaScript1.2">
<!--
function confirm_delete(){
	if (confirm("��ȷ�����Ҫ��ô����")){
		return true;
	}
	return false;
}
//-->
</script>


<!-- �ж��Ƿ������� -->
 <%
 if rs.eof and rs.bof then
 %>
 
<table width="400" border="0" cellpadding="3" cellspacing="1" class="a2">
  <tr>
    <td><div align="center">��ʱ��û�����ԣ�</div></td>
  </tr>
</table>
<%
response.end
end if
%>



<!-- ��ҳ���ܴ����,�ɶ���ʹ�� -->
<% 
 if not rs.eof then
  rs.MoveFirst  'ע��ŵ�ǰ����,�����κ�ҳ�����ڵ�һ����¼��
  end if
  rs.pagesize=MaxPerPage  '����ÿҳ�����ʾ��������¼
  If trim(Request("Page"))<>"" then  '��������ҳ�β�Ϊ��
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
 
  <table width="95%" border="0" cellpadding="3" cellspacing="1" bgcolor="#DEDFDE" class="a2">
    <tr bgcolor="#F7F7F7" class="a1">
      <td width="214"><div align="center">��<%=id%>λ</div></td>
      <td><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0" class="a1">
        <tr>
          <td width="98">����ʱ�䣺</td>
          <td width="510"><%=rs("time")%></td>

    
          <td width="29"><div align="center"><img src="../images/friend.gif" alt="���ԣ�<%=rs("where")%>" width="18" height="18"></div></td>
          <td width="30"><div align="center"><img src="../images/ip.gif" alt="IP��<%=rs("ip")%>" width="18" height="18"></div></td>
        </tr>
      </table></td>
    </tr>
    <tr class="a3">
      <td valign="top" bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>��</td>
        </tr>
        <tr>
          <td><div align="center"><img src="../<%=rs("pic")%>" width="75" height="75"></div></td>
        </tr>
        <tr>
          <td height="35">
            <div align="center"><strong><%=rs("name")%></strong></div></td>
        </tr>
      </table>

      <div align="center">
        <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><div align="center">[<span onClick='return confirm_delete()'><a href="?i=2&id=<%=rs("id")%>">ɾ��</a></span>]��[<a href="?i=3&id=<%=rs("id")%>">�ظ�</a>]</div></td>
          </tr>
        </table>
		
      </div></td>
      <td valign="top" class="a4"><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" style="table-layout:fixed;word-break:break-all">
        <tr>
          <td width="7%"></td>
          <td width="93%">              <%if rs("show")=2 then
response.write"����������Ϊ���Ļ���ֻ�й���Ա�ſɲ鿴!<br><br>"
End If
%><%=rs("title")%></td>
        </tr>
        <tr>
          <td colspan="2"><%=rs("content")%>
          </td>
        </tr>
      </table>
	  
	  <%
		if rs("reply")<>"" then%>
        <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" style="table-layout:fixed;word-break:break-all">
          <tr>
            <td><div align="center">
              <hr width="98%" size="1" noshade color="#cccccc">
             </div></td>
          </tr>
          <tr>
            <td><font color="#999999">վ���ظ���<%=htmlencode(rs("reply"))%></font></td>
          </tr>
      </table><%end if%>      </td>
    </tr>
</table>
  
	  
	  
	  </td>
    </tr>
  </table>
  <br>
  <% i=i+1
  rs.movenext
  loop
  %>
  <br>
  <table width="95%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><!-- ��ҳ��ʾ����� �ɶ���ʹ��,ע��������ҳ���ܴ�����ʹ�� -->
        <div align="center">��ǰ��<font color="#ff0000"><%=currentpage%></font>ҳ �ܹ�<font color="#FF0000"><%=n%></font>ҳ ��<font color="#FF0000"><%=rs.recordcount%></font>������ 
          <%k=currentPage
		if k<>1 then
			response.write "[<b>"+"<a href='?page=1'>��ҳ</a></b>] "
			response.write "[<b>"+"<a href='?page="&cstr(k-1)&"'>��һҳ</a></b>] "
		else
			Response.Write "<s>[��ҳ] [��һҳ]</s>"
		end if
		if k<>n then
			response.write "[<b>"+"<a href='?page="&cstr(k+1)&"'>��һҳ</a></b>] "
			response.write "[<b>"+"<a href='?page="&cstr(n)&"'>βҳ</a></b>] "
		else
			Response.Write "<s>[��һҳ] [βҳ]</s>"
		end if
          
	%></td>
    </tr>
</table>
<%
rs.close             
%>

<%
end sub
%>

<%
'------------------------
'�༭����
'------------------------
sub b()
%>
<%
Dim id
id=trim(request("id"))
If IsNumeric(id) = False Then
	GoError "��ͨ��ҳ���ϵ����ӽ��в�������Ҫ��ͼ�ƻ���ϵͳ��"
End If

sSql="delete * from book where id="&id
Conn.execute sSql
 Response.Write "<SCRIPT LANGUAGE='JavaScript'>"
 Response.Write "window.location.href='?'"
 Response.Write "</SCRIPT>"

%>
<%
end sub
%>


<%
'------------------------
'�༭����
'------------------------
sub c()
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">

<SCRIPT language=JavaScript>

function CheckInput(){

	if(form.reply.value==''){
		alert("���ݲ���Ϊ�գ�");
		form.reply.focus();
		return false;
	}
	if(form.reply.value.length>1000){
		alert("���ݲ��ܳ���1000���ַ���");
		form.reply.focus();
		return false;
	}
	
	return true;
}
</SCRIPT>
<%
Dim id
id=trim(request("id"))
If IsNumeric(id) = False Then
	GoError "��ͨ��ҳ���ϵ����ӽ��в�������Ҫ��ͼ�ƻ���ϵͳ��"
End If
set oRs=Server.CreateObject("ADODB.Recordset")

sSql="select * from book where id="&id
oRs.open sSql,Conn,1,3
%>


<%
if request("hiddenField")=1 then
dim sql,BookReply,i

For i = 1 To Request.Form("reply").Count
	BookReply = BookReply & Request.Form("reply")(i)
Next

BookReply=CheckStr(BookReply)
sql="Update book set reply='"&BookReply&"' Where id="&id
conn.execute sql
response.write "<p align=center>���Իظ��ɹ���<script>window.setTimeout(""location.href='?'"",1000);</script></p>"
Response.End
end if
%>


<table width="95%" border="0" cellpadding="3" cellspacing="1" bgcolor="#DEDFDE" class="a2">
  <form action="?i=3&action=save&id=<%=request("id")%>" method="post" name="form" id="form" onsubmit=return(CheckInput())><tr class="a1">
    <td colspan="2"><div align="center"><strong>�� �� Ա �� ��</strong></div></td>
    </tr>
    <tr class="a3">
      <td colspan="2"><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td>�ظ� [ <font color="#ff0000"><%=oRs("name")%></font> ] ������</td>
          </tr>
            </table></td>
    </tr>
    <tr class="a3">
      <td width="18%">
		<p align="right">�ظ����ݣ�
      <input name="hiddenField" type="hidden" value="1"></td>
      <td width="80%"><textarea name="reply" cols="90" rows="10" id="reply" class="input_text"><%=oRs("reply")%></textarea></td>
    </tr>
    <tr class="a3">
      <td height="40" colspan="2"><div align="center">
        <input type="submit" name="Submit" value="�ظ�����" class="input_submit">
��
<input type="reset" name="Submit2" value="�����д" class="input_submit">
��
<input type="submit" name="Submit3" value="��������" class="input_submit"> </div></td>
    </tr>
  </form>
</table>
<%
oRs.Close
%>
<%
end sub
%>




