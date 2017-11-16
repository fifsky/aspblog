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
 Response.Write "alert('留言昵称不能为空！');"
 Response.Write "history.go(-1);"
 Response.Write "</SCRIPT>"
	Response.End
end if
if bookcontent="" then

 Response.Write "<SCRIPT LANGUAGE='JavaScript'>"
 Response.Write "alert('留言内容不能为空！');"
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
 Response.Write "alert('恭喜您！您已经成功提交了信息！');"
 Response.Write "window.location.href='gbook.asp?i=show'"
 Response.Write "</SCRIPT>"
Response.end
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>留言 - <%=WebName%></title>
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
   const MaxPerPage=6   '单独页最大记录数 const 用来申明常量
   dim sql 
   dim rs
   dim totalPut   '总记录
   dim CurrentPage  '当前页次
   dim TotalPages   '总页数
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
                                <td align="center">暂时还没有任何留言！</td>
                              </tr>
                            </table>
                          <%
else
%>
                            <!-- 分页功能代码块,可独立使用 -->
                            <% 
 if not rs.eof then
  rs.MoveFirst  '注意放到前面来,否则到任何页总是在第一个记录上
  end if
  rs.pagesize=MaxPerPage  '设置每页最多显示多少条记录
  If trim(Request("Page"))<>"" then  '如果请求的页次不为空
   If IsNumeric(trim(Request("Page"))) = False Then
		GoError "分页参数错误，请不要试图破坏此系统。"
	End If
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
                            <table width="89%" border="0" align="center" cellpadding="3" cellspacing="1" class="alpha">
                              <tr>
                                <td width="75" align="center">【第<%=id%>位】</td>
                                <td width="376"><table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
                                    <tr>
                                      <td width="57">发表于：</td>
                                      <td align="left"><%=rs("time")%></td>
                                      <td width="26" align="center"><img src="images/friend.gif" alt="来自：<%=rs("where")%>" width="18" height="18" /></td>
                                      <td width="30" align="center"><img src="images/ip.gif" alt="IP：<%=rs("ip")%>" width="18" height="18" /></td>
                                    </tr>
                                </table></td>
                              </tr>
                              <tr bgcolor="#666666">
                                <td height="1" colspan="3" ></td>
                              </tr>
                              <tr>
                                <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td>　</td>
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
response.write"此留言内容为悄悄话，只有管理员才可查看!"
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
                                        <td align="left"><span class="style3">------以下是本站站长回复的内容------</span></td>
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
                                <td align="center"> 第 <font color="#FF0000"><%=currentpage%>/<%=n%></font> 页 共 <font color="#FF0000"><%=rs.recordcount%></font> 条留言
                                  <%k=currentPage
		if k<>1 then
			response.write "[<b>"+"<a href='?action=show&page=1'>首页</a></b>] "
			response.write "[<b>"+"<a href='?action=show&page="&cstr(k-1)&"'>上一页</a></b>] "
		else
			Response.Write "[首页] [上一页]"
		end if
		if k<>n then
			response.write "[<b>"+"<a href='?action=show&page="&cstr(k+1)&"'>下一页</a></b>] "
			response.write "[<b>"+"<a href='?action=show&page="&cstr(n)&"'>尾页</a></b>] "
		else
			Response.Write "[下一页] [尾页]"
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
                                  <td width="79" height="23" align="right">留言昵称：</td>
                                  <td colspan="3" align="left"><input name="name" type="text" size="15" class="input2" />
                                    * （请使用中文）</td>
                                  <td width="194" rowspan="2" align="left"><img src="images/face/1.gif" name="iface" width="75" height="75" id="iface" /></td>
                                </tr>
                                <tr>
                                  <td height="26" align="right">来自哪里：</td>
                                  <td width="71" align="left"><select name="where" size="1" class="input2" id="select2" style="border: 1px solid #E6E6E6">
                                      <option value="未知" selected="selected">保密</option>
                                      <option value="北京">北京</option>
                                      <option value="上海">上海</option>
                                      <option value="天津">天津</option>
                                      <option value="重庆">重庆</option>
                                      <option value="河北">河北</option>
                                      <option value="山西">山西</option>
                                      <option value="内蒙古">内蒙古</option>
                                      <option value="辽宁">辽宁</option>
                                      <option value="吉林">吉林</option>
                                      <option value="黑龙江">黑龙江</option>
                                      <option value="江苏">江苏</option>
                                      <option value="浙江">浙江</option>
                                      <option value="安徽">安徽</option>
                                      <option value="福建">福建</option>
                                      <option value="江西">江西</option>
                                      <option value="山东">山东</option>
                                      <option value="河南">河南</option>
                                      <option value="湖北">湖北</option>
                                      <option value="湖南">湖南</option>
                                      <option value="广东">广东</option>
                                      <option value="广西">广西</option>
                                      <option value="海南">海南</option>
                                      <option value="四川">四川</option>
                                      <option value="贵州">贵州</option>
                                      <option value="云南">云南</option>
                                      <option value="西藏">西藏</option>
                                      <option value="陕西">陕西</option>
                                      <option value="甘肃">甘肃</option>
                                      <option value="宁夏">宁夏</option>
                                      <option value="青海">青海</option>
                                      <option value="新疆">新疆</option>
                                      <option value="香港">香港</option>
                                      <option value="澳门">澳门</option>
                                      <option value="台湾">台湾</option>
                                      <option value="海外">海外</option>
                                    </select>
                                  </td>
                                  <td width="66" align="left">选择头像：</td>
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
                                  <td height="26" align="right">留言性质：</td>
                                  <td colspan="4" align="left"><input name="show" type="radio" class="input2" value="1" checked="checked" />
                                    公开
                                    <input name="show" type="radio" class="input2" value="2" />
                                    悄悄话</td>
                                </tr>
                                <tr>
                                  <td align="right" valign="middle">留言内容：<br />
                                      <br />
                                  </td>
                                  <td colspan="4" align="left"><textarea name="content" cols="40" rows="4" id="textarea3" class="input2"></textarea>
                                  </td>
                                </tr>
                                <tr>
                                  <td colspan="5" valign="middle" align="center"><input name="submit" type="submit" class="input2" id="submit3" value="发表留言" onClick="javascript:return CheckForm();"/>
                                      <input type="reset" name="Submit2" value="清除留言" class="input2" />
                                      <input type="button" value="返　　回" name="cmdExit" onClick=" history.back()" class="input2" />
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
                  <TD align="right" VALIGN="TOP" ><IMG SRC="images/up.gif" WIDTH="50" HEIGHT="20" onMouseOver="scroll(-1)" onMouseOut="scroll(0)" onMouseDown="scroll(-3)" BORDER="0" ALT="按下鼠标速度会更快！"> <IMG SRC="images/down.gif" onMouseOver="scroll(1)" onMouseOut="scroll(0)"  onmousedown="scroll(3)" BORDER="0" WIDTH="50" HEIGHT="20" ALT="按下鼠标速度会更快！"> </TD>
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