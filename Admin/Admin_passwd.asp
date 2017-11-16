<!--#include file="Error.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<link href=images/admin.css rel=stylesheet>
<META http-equiv=Content-Type content=text/html;charset=GB2312>
<%set rs=server.createobject("adodb.recordset")%>
<%
if request("action")="change" then
pass=request("admin_pass")
if pass="" then
 Response.Write "<SCRIPT LANGUAGE='JavaScript'>"
 Response.Write "alert('密码不能为空！');"
 Response.Write "history.go(-1);"
 Response.Write "</SCRIPT>"
 Response.End
end if

sql="select top 1 * from admin"
rs.open sql,conn,1,3
rs("admin_name")=request("admin_name")
rs("admin_pass")=md5(request("admin_pass"))
rs.update
rs.close
response.write"<SCRIPT language=JavaScript>alert('管理员信息修改成功！');"
response.write"location.href='?'</SCRIPT>"
else
%>

<SCRIPT language=javascript id=clientEventHandlersJS>
<!--
function form1_onsubmit()
{
if (document.FORM1.admin_pass.value!=document.FORM1.admin_pass0.value)
{
alert ("两次密码输入不一致！");
document.FORM1.admin_pass.value='';
document.FORM1.admin_pass0.value='';
document.FORM1.admin_pass.focus();
return false;
}
}
//-->
</SCRIPT>
<body Style="background-color:#9EB6D8" text="#000000" leftmargin="5" topmargin="5">
<form name="FORM1" method="POST" action="?action=change" onSubmit="return form1_onsubmit()" >

                      <%sql="select top 1 * from admin"
						  rs.open sql,conn,3,3%>

	
  <table width="400" border="0" cellpadding="3" cellspacing="1" class="a2">
	 <tr bgcolor="#F7F7F7">
       <td colspan="2" class="a1"><div align="center"><strong>管理员信息设置</strong></div></td>
    </tr>
       <tr bgcolor="#FFFFFF" class="a3">
         <td><div align="center">管理员名称</div></td>
         <td><div align="center">
           <input type="text" name="admin_name" size="20" value="<%=rs("admin_name")%>" class=input >
         </div></td>
    </tr>
       <tr bgcolor="#FFFFFF" class="a4">
         <td><div align="center">管理员密码</div></td>
         <td><div align="center">
           <input type="password" name="admin_pass" size="20" class=input>
         </div></td>
    </tr>
       <tr bgcolor="#FFFFFF" class="a3">
         <td><div align="center">密码确认</div></td>
         <td><div align="center">
           <input type="password" name="admin_pass0" size="20" class=input>
         </div></td>
    </tr>
       <tr bgcolor="#FFFFFF">
         <td colspan="2" class="a4"><div align="center">
           <input type="submit" name="Submit2" value="确定修改" class=input>
         </div></td>
       </tr>

  </table>
</form>

       <%rs.close%>
       <%end if
set rs=nothing
conn.close
set conn=nothing
%>
                 
     