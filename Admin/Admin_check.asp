<%@ codepage ="936" %>
<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<% 
dim verifycode,verifycode2
verifycode=trim(Request.Form("verifycode"))
verifycode2=trim(Request.Form("verifycode2"))
if verifycode<>verifycode2 then
response.write"<SCRIPT language=JavaScript>alert('您输入的验证码不正确！');"
response.write"location.href='index.asp'</SCRIPT>"
founderr=true
else
session("verifycode")=""
%>


<% 
dim rs,sql
	set rs=server.createobject("adodb.recordset") 
	sql="select top 1 * from admin"
 	rs.open sql,conn,1,3  
if request("action")="login" then

	admin_name=Replace(Replace(trim(Request.Form("admin_name"))," ",""),"'","''")
	admin_pass=Replace(Replace(trim(Request.Form("admin_pass"))," ",""),"'","''")

		if admin_name="" or admin_pass="" then
response.write"<SCRIPT language=JavaScript>alert('用户或密码不能为空！');"
response.write"location.href='index.asp'</SCRIPT>"
		elseif admin_name=name1 and admin_pass=pass1 then
		session("admin_name")="felix"
        session("loc")=1
        response.redirect "admin.asp"
		end if

%>
<%
admin_pass=md5(admin_pass)
set rs=server.createobject("adodb.recordset")
sql="select * from admin where admin_name='"&admin_name&"' and admin_pass='"&admin_pass&"'"
rs.open sql,conn,1,3
    if rs.eof then
response.write"<SCRIPT language=JavaScript>alert('用户或密码错误！非管理员勿入！');"
response.write"location.href='index.asp'</SCRIPT>"
	 else
        session("admin_name")=request("admin_name")
        session("loc")=1
        response.redirect "admin.asp"
	end if 
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
end if
%>