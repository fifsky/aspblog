<!--#include file="conn.asp"-->
<% dim rndnum,verifycode
Randomize
Do While Len(rndnum)<4
num1=CStr(Chr((57-48)*rnd+48))
rndnum=rndnum&num1
loop
session("verifycode")=rndnum
%>
<% 
dim rs,sql
	set rs=server.createobject("adodb.recordset") 
	sql="select top 1 * from admin"
 	rs.open sql,conn,1,3  
%>
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=rs("mingcheng")%></title>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
	color: #FFFFFF;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
.style1 {font-size: 12px}
.style2 {color: #FF0000}
body {
	background-color: #232323;
}
-->
</style></head>
<body>
<br>
<br>
<br>
<br>
<table width="413" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#232323">
  <tr> 
    <td height="29" align="center" bgcolor="#232323">&nbsp;</td>
  </tr>
  <tr> 
    <td height="25" align="center"><img src="images/vtsj.gif" width="125" height="83"></td>
  </tr>
  <tr bgcolor="#232323"> 
    <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td bgcolor="#232323"> <form action="admin_check.asp?action=login" method="POST">
              <table width="95%" height="142" border="0" align="center">
                <tr> 
                  <td> <fieldset>
                    <legend align="left" class="style1" accesskey="F">登陆窗口</legend>
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
                      <tr> 
                        <td width="10%" height="23" bgcolor="#232323"><font size="2">&nbsp;</font></td>
                        <td width="20%" bgcolor="#232323"><span class="style1">用&nbsp;户&nbsp;名：</span></td>
                        <td bgcolor="#232323"><input name="admin_name" type="text" id="admin_name2" style="border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onFocus="this.select(); " onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'">
                        </td>
                      </tr>
                      <tr> 
                        <td width="10%" height="25" bgcolor="#232323"><font size="2">&nbsp;</font></td>
                        <td width="20%" bgcolor="#232323"><span class="style1">密&nbsp;&nbsp;&nbsp;&nbsp;码：</span></td>
                        <td bgcolor="#232323"><input name="admin_pass" type="password" id="admin_pass2" style="border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onFocus="this.select(); " onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'"> 
                        </td>
                      </tr>
                      <tr> 
                        <td align="center" bgcolor="#232323">&nbsp;</td>
                        <td bgcolor="#232323"><span class="style1">验&nbsp;证&nbsp;码:</span></td>
                        <td bgcolor="#232323"><input name="verifycode" id="verifycode" style="border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onFocus="this.select(); " onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" size="6" maxlength="4"> 
                          <span class="style1"><font color="#000000">请在左边输入　</font><span class="style2"><%=session("verifycode")%> 
                                </span><font color="#000000">
                                <input type="hidden" name="verifycode2" value="<%=session("verifycode")%>">
                        </font>&nbsp;</font></span></td>
                      </tr>
                      <tr> 
                        <td colspan="2" align="center" bgcolor="#232323">&nbsp;</td>
                        <td bgcolor="#232323"> <input name="submit" type="submit" class="input" value=" 登 陆 "> 
                        <input type="reset" name="submit2" value=" 重 填 " class="input"></td>
                      </tr>
                    </table>
                    </fieldset></td>
                </tr>
              </table>
          </form></td>
        </tr>
    </table></td>
  </tr>
</table>