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
   const MaxPerPage=5   '单独页最大记录数 const 用来申明常量
   dim sql 
   dim rs
   dim totalPut   '总记录
   dim CurrentPage  '当前页次
   dim TotalPages   '总页数
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
	if (confirm("你确定真的要这么做吗？")){
		return true;
	}
	return false;
}
//-->
</script>


<!-- 判断是否有留言 -->
 <%
 if rs.eof and rs.bof then
 %>
 
<table width="400" border="0" cellpadding="3" cellspacing="1" class="a2">
  <tr>
    <td><div align="center">暂时还没有留言！</div></td>
  </tr>
</table>
<%
response.end
end if
%>



<!-- 分页功能代码块,可独立使用 -->
<% 
 if not rs.eof then
  rs.MoveFirst  '注意放到前面来,否则到任何页总是在第一个记录上
  end if
  rs.pagesize=MaxPerPage  '设置每页最多显示多少条记录
  If trim(Request("Page"))<>"" then  '如果请求的页次不为空
	CurrentPage= CLng(request("Page"))   'clng是转换成长整型数据类型,并赋值到当前页次上
	If CurrentPage> rs.PageCount then  '如果当前页次大于总页数,则将最大页次赋值到当前页次上
		CurrentPage = rs.PageCount 
	End If 
Else 
	CurrentPage= 1 '一切条件不成立,将当前页设为第一页
End If 


 totalPut=rs.recordcount '将总记录赋值于TOTALPUT
	if CurrentPage<>1 then '如果当前页数不等于第一页
		if (currentPage-1)*MaxPerPage<totalPut then  '如果当前页减一乘以每页最大的记录数小于总记录的话
			rs.move(currentPage-1)*MaxPerPage  '相对当前记录数向后移动
			dim bookmark  '定义书签变量
			bookmark=rs.bookmark '将当前记录的标签赋于变量BOOKMARK上
		end if 
	end if

	dim n,k 
	if (totalPut mod MaxPerPage)=0 then  '总记录数与每页最大记录数求余的结果为零时,则N返回整数页次,否则再加一.
		n= totalPut \ MaxPerPage
	else  
		n= totalPut \ MaxPerPage + 1  
	end if
%>
 



<!-- 将RS记录指针指向第一个记录，然后开始判断移动记录时，记录结尾是否为空，如果不为空接着移动指针，把所有数据都读取出来。直到结尾为空时，退出循环 -->
 <%
 id=(totalPut-MaxPerPage*(currentPage-1))+1 
i=0
Do While Not rs.EOF and i<MaxPerPage
id=id-1
%>
 
  <table width="95%" border="0" cellpadding="3" cellspacing="1" bgcolor="#DEDFDE" class="a2">
    <tr bgcolor="#F7F7F7" class="a1">
      <td width="214"><div align="center">第<%=id%>位</div></td>
      <td><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0" class="a1">
        <tr>
          <td width="98">发表时间：</td>
          <td width="510"><%=rs("time")%></td>

    
          <td width="29"><div align="center"><img src="../images/friend.gif" alt="来自：<%=rs("where")%>" width="18" height="18"></div></td>
          <td width="30"><div align="center"><img src="../images/ip.gif" alt="IP：<%=rs("ip")%>" width="18" height="18"></div></td>
        </tr>
      </table></td>
    </tr>
    <tr class="a3">
      <td valign="top" bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>　</td>
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
            <td><div align="center">[<span onClick='return confirm_delete()'><a href="?i=2&id=<%=rs("id")%>">删除</a></span>]　[<a href="?i=3&id=<%=rs("id")%>">回复</a>]</div></td>
          </tr>
        </table>
		
      </div></td>
      <td valign="top" class="a4"><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" style="table-layout:fixed;word-break:break-all">
        <tr>
          <td width="7%"></td>
          <td width="93%">              <%if rs("show")=2 then
response.write"此留言内容为悄悄话，只有管理员才可查看!<br><br>"
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
            <td><font color="#999999">站长回复：<%=htmlencode(rs("reply"))%></font></td>
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
      <td><!-- 分页显示代码块 可独立使用,注意和上面分页功能代码配使用 -->
        <div align="center">当前第<font color="#ff0000"><%=currentpage%></font>页 总共<font color="#FF0000"><%=n%></font>页 共<font color="#FF0000"><%=rs.recordcount%></font>个留言 
          <%k=currentPage
		if k<>1 then
			response.write "[<b>"+"<a href='?page=1'>首页</a></b>] "
			response.write "[<b>"+"<a href='?page="&cstr(k-1)&"'>上一页</a></b>] "
		else
			Response.Write "<s>[首页] [上一页]</s>"
		end if
		if k<>n then
			response.write "[<b>"+"<a href='?page="&cstr(k+1)&"'>下一页</a></b>] "
			response.write "[<b>"+"<a href='?page="&cstr(n)&"'>尾页</a></b>] "
		else
			Response.Write "<s>[下一页] [尾页]</s>"
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
'编辑新闻
'------------------------
sub b()
%>
<%
Dim id
id=trim(request("id"))
If IsNumeric(id) = False Then
	GoError "请通过页面上的链接进行操作，不要试图破坏此系统。"
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
'编辑新闻
'------------------------
sub c()
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">

<SCRIPT language=JavaScript>

function CheckInput(){

	if(form.reply.value==''){
		alert("内容不能为空！");
		form.reply.focus();
		return false;
	}
	if(form.reply.value.length>1000){
		alert("内容不能超过1000个字符！");
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
	GoError "请通过页面上的链接进行操作，不要试图破坏此系统。"
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
response.write "<p align=center>留言回复成功！<script>window.setTimeout(""location.href='?'"",1000);</script></p>"
Response.End
end if
%>


<table width="95%" border="0" cellpadding="3" cellspacing="1" bgcolor="#DEDFDE" class="a2">
  <form action="?i=3&action=save&id=<%=request("id")%>" method="post" name="form" id="form" onsubmit=return(CheckInput())><tr class="a1">
    <td colspan="2"><div align="center"><strong>管 理 员 回 复</strong></div></td>
    </tr>
    <tr class="a3">
      <td colspan="2"><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td>回复 [ <font color="#ff0000"><%=oRs("name")%></font> ] 的留言</td>
          </tr>
            </table></td>
    </tr>
    <tr class="a3">
      <td width="18%">
		<p align="right">回复内容：
      <input name="hiddenField" type="hidden" value="1"></td>
      <td width="80%"><textarea name="reply" cols="90" rows="10" id="reply" class="input_text"><%=oRs("reply")%></textarea></td>
    </tr>
    <tr class="a3">
      <td height="40" colspan="2"><div align="center">
        <input type="submit" name="Submit" value="回复留言" class="input_submit">
　
<input type="reset" name="Submit2" value="清除重写" class="input_submit">
　
<input type="submit" name="Submit3" value="返　　回" class="input_submit"> </div></td>
    </tr>
  </form>
</table>
<%
oRs.Close
%>
<%
end sub
%>




