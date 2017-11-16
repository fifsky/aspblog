<%@ LANGUAGE="VBScript" %>
<!--#include file="conn.asp"-->
<!--#include file="Error.asp"-->
<%
if request("action")="del" then
set rs=server.createobject("adodb.recordset")
sql="delete from feedback where ID="&request("id")
rs.open sql,conn,1,1
set rs=nothing
conn.close
set conn=nothing
response.redirect "?"
else
%>
<link href=images/admin.css rel=stylesheet>
<META http-equiv=Content-Type content=text/html;charset=GB2312>
<script language="JavaScript">
function ConfirmDel()
{
if (confirm("你真的要删除此信息吗!"))
	return true;
else
	return false;
}
</script>
<body Style="background-color:#9EB6D8" text="#000000" leftmargin="5" topmargin="5">

      <table width="550" border="0" cellpadding="3" cellspacing="1" class="a2">
        <tr class="a1"> 
          <td height="20"> <div align="center">反馈单管理
              </div></td>
        </tr>
        <tr> 
			<td><div align="center"> 
			<% 
const MaxPerPage=5 '分页显示的纪录个数 
dim sql 
dim rs 
dim totalPut 
dim CurrentPage 
dim TotalPages 
dim i,j 
 
 
if not isempty(request("page")) then 
currentPage=cint(request("page")) 
else 
currentPage=1 
end if 
set rs=server.createobject("adodb.recordset") 
sql="select * from feedback order by id desc" 
rs.open sql,conn,1,1 
 
if rs.eof and rs.bof then 
response.write "<p align='center'>还没有任何反馈!</p>" 
else 
totalPut=rs.recordcount 
totalPut=rs.recordcount 
if currentpage<1 then 
currentpage=1 
end if 
if (currentpage-1)*MaxPerPage>totalput then 
if (totalPut mod MaxPerPage)=0 then 
currentpage= totalPut \ MaxPerPage 
else 
currentpage= totalPut \ MaxPerPage + 1 
end if 
end if 
if currentPage=1 then 
showpages 
showContent 
showpages 
else 
if (currentPage-1)*MaxPerPage<totalPut then 
rs.move (currentPage-1)*MaxPerPage 
dim bookmark 
bookmark=rs.bookmark 
showpages 
showContent 
showpages 
else 
currentPage=1 
showContent 
end if 
end if 
rs.close 
end if 
set rs=nothing 
conn.close 
set conn=nothing 
sub showContent 
dim i 
i=0 
do while not (rs.eof or err) %>
            
                <table width="100%" border="0" cellpadding="3" cellspacing="1" class="a2">
                <tr class="a3"> 
                  <td width="16%" height="25"> <div align="center">订购产品</div></td>
                  <td colspan="2">&nbsp;<b><%=rs("jobname")%></b> 
                  &nbsp; </td>
                  <td><div align="center"><a href="?action=del&id=<%=rs("ID")%>" 
onclick="return ConfirmDel()">[ 删 除 ] </a></div></td>
                </tr>
                <%if rs("email")<>"" then%>
                <tr class="a4"> 
                  <td height="25"> <div align="center">公司名称</div></td>
                  <td width="41%">&nbsp;<%=rs("school")%>&nbsp;</td>
                  <td width="13%"><div align="center"><font color="#FF0000">订购日期</font></div></td>
                  <td width="30%"><font color="#FF0000"><%=rs("time")%></font></td>
                </tr>
                <tr class="a3"> 
                  <td height="25"> <div align="center">联系电话</div></td>
                  <td>&nbsp;<%=rs("telephone")%></td>
                  <td><div align="center">联系人</div></td>
                  <td><%=rs("mane")%></td>
                </tr>
                <tr class="a4"> 
                  <td height="25"><div align="center">联系地址</div></td>
                  <td colspan="3">&nbsp;<%=rs("address")%></td>
                </tr>
                <tr class="a3"> 
                  <td height="25"><div align="center">E-mail</div></td>
                  <td>&nbsp;<%=rs("email")%></td>
                  <td><div align="center">QQ号</div></td>
                  <td><%=rs("specialty")%></td>
                </tr>
               
                <%end if%>
                <tr class="a4"> 
                  <td height="25"><div align="center">备注</div></td>
                  <td height="25" colspan="3">&nbsp;<%=rs("ability")%></td>
                </tr>
                
              </table>
<br>

			  <table width="550" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><div align="center"><% i=i+1 
if i>=MaxPerPage then exit do 
rs.movenext 
loop 
end sub 
sub showpages() 
dim n 
if (totalPut mod MaxPerPage)=0 then 
n= totalPut \ MaxPerPage 
else 
n= totalPut \ MaxPerPage + 1 
end if 

dim k 
response.write "<p align='left'>&gt;&gt; 分页 " 
for k=1 to n 
if k=currentPage then 
response.write "[<b>"+Cstr(k)+"</b>] " 
else 
response.write "[<b>"+"<a href='?page="+cstr(k)+"'>"+Cstr(k)+"</a></b>] " 
end if 
next 
end sub 
%></div></td>
        </tr>
      </table>

              </div></td>
        </tr>
</table>


<%end if%>