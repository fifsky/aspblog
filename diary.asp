<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="include.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>日记 - <%=WebName%></title>
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
<div id="Layer1" style="position:absolute; left:345px; top:0px; width:305px; height:84px; z-index:1">
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="333" height="80">
    <param name="movie" value="images/VTsj.swf">
    <param name="quality" value="high">
    <param name="wmode" value="transparent">
    <param name="menu" value="false">
    <embed src="images/VTsj.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="333" height="80"></embed>
  </object>
</div>
<P> </td>
	  <td colspan="2">
			<img src="images/VTsj_16.jpg" width="68" height="171" alt=""></td>
	  <td width="427" height="414" colspan="5" rowspan="2" valign="top" background="images/VTsj_17.jpg"><table width="100%"  border="0" cellpadding="5" cellspacing="0">
        <tr>
          <td width="416"><table width="100%"  border="0" cellpadding="3" cellspacing="0">
              <tr>
                <td height="35"><DIV class="site-menu2">
                    <%
set rsa=server.createobject("adodb.recordset") 
 sqla="select * from HN_newscla where MidCls<>'' order by id ASC" 
 rsa.open sqla,conn,1,1
DO WHILE Not rsa.eof
%>
                    <li><a href=diary.asp?id=<%=rsa("id")%>><%=RSa("MidCls")%></a></li>
                  <%
rsa.movenext
LOOP
rsa.close
set rsa=nothing
%>
                </DIV></td>
              </tr>
              <%
cid=request("id")
if cid="" then
sql="select * from HN_news where kinds=true order by id desc"
else
sql="select * from HN_news where cstr(class)='"&cid&"' and kinds=true order by id desc"
end if
set rsnews=server.createobject("adodb.recordset")
rsnews.open sql,conn,1,1
	if rsnews.eof then 
	Response.Write "</td></tr></table>"
	end if
IF Not rsnews.eof Then
proCount=rsnews.recordcount
	rsnews.PageSize=10
		if not IsEmpty(Request("ToPage")) then
	       ToPage=CInt(Request("ToPage"))
		   if ToPage>rsnews.PageCount then
					rsnews.AbsolutePage=rsnews.PageCount
					intCurPage=rsnews.PageCount
		   elseif ToPage<=0 then
					rsnews.AbsolutePage=1
					intCurPage=1
				else
					rsnews.AbsolutePage=ToPage
					intCurPage=ToPage
				end if
			else
			        rsnews.AbsolutePage=1
					intCurPage=1
		  end if
	intCurPage=CInt(intCurPage)
	For i = 1 to rsnews.PageSize
	if rsnews.EOF then     
	Exit For 
	end if
%>
              <tr>
                <td width="90%" height=12 background="images/bg_line.gif">·<a href="View.asp?id=<%=rsnews("id")%>" target="_blank" class="photo">
                  <% = Left(rsnews("title"),30) %>
                  <% if Len(rsnews("title"))>30 then %>
                  …
                  <% end if %>
                </a><span class="style4">(<%=rsnews("time")%>)</span></td>
              </tr>
              <%
 rsnews.MoveNext 
 next
%>
            </table>
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td width="51%" height="30" align="right"><%if intCurPage<>1 then%>
                      <a href="?id=<%=cid%>&amp;ToPage=1" class="photo">首页</a> <a href="?id=<%=cid%>&amp;ToPage=<%=intCurPage-1%>" class="photo">上一页</a>
                      <% end if
if intCurPage<>rsnews.PageCount then %>
                      <a href="?id=<%=cid%>&amp;ToPage=<%=intCurPage+1%>" class="photo">下一页</a> <a href="?id=<%=cid%>&amp;ToPage=<%=rsnews.PageCount%>" class="photo">尾页</a>
                      <% end if%>                  </td>
                </tr>
              </table>
            <%end if%>
              <%
rsnews.close
set rsnews=nothing
%>          </td>
        </tr>
      </table></td>
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