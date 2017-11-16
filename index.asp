
<!--#include file="conn.asp"-->
<!--#include file="include.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=WebName%> - 首页</title>
<meta name="keywords" content="<%=W_WebSiteKeyword%>">
<LINK href="css.css" type=text/css rel=stylesheet>
<link rel="stylesheet" href="inc/lightbox.css" type="text/css" media="screen" />
<SCRIPT language="JavaScript" src="inc/sampleplay2.js"></SCRIPT>
<script src="inc/prototype.js" type="text/javascript"></script>
<script src="inc/scriptaculous.js?load=effects" type="text/javascript"></script>
<script src="inc/lightbox.js" type="text/javascript"></script>
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
      </div></td>
	  <td colspan="2">
			<img src="images/VTsj_16.jpg" width="68" height="171" alt=""></td>
	  <td width="427" height="414" colspan="5" rowspan="2" align="center" valign="top" background="images/VTsj_17.jpg">
<table width="427"  border="0" cellpadding="5" cellspacing="0">
  <tr>
    <td width="515"><img src="images/idiary.gif" alt="" width="414" height="29" border="0" usemap="#MapMap">
        <map name="MapMap">
          <area shape="rect" coords="48,3,130,26" href="diary.asp">
        </map>
        <table width="100%"  border="0" cellpadding="3" cellspacing="0">
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
	rsnews.PageSize=5
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
        <%end if%></td></tr>
</table>
<img src="images/imedia.gif" alt="" width="414" height="29" border="0" usemap="#MapMapMap">
<map name="MapMapMap">
  <area shape="rect" coords="48,4,133,26" href="media.asp">
</map>
<br>
<table width="427"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50" align="center" valign="top"><table width="100%"  border="0" cellpadding="3" cellspacing="0">
      <tr>
        <td width="395" height="32"><table align="center">
          <tr>
            <td width="22"><%
dim ii
ii=0
cid=request("id")
if cid="" then
sql="select * from HN_down where kinds=true order by id desc"
else
sql="select * from HN_down where cstr(class)='"&cid&"' and kinds=true order by id desc"
end if
set rsnews=server.createobject("adodb.recordset")
rsnews.open sql,conn,1,1
	if rsnews.eof then 
	Response.Write "</td></tr></table></td></tr></table>"
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

ii=ii+1
if ii mod 2=1 then response.Write("<tr>")
%>
            <td width="172" height=12 class="style4"><% = Left(rsnews("title"),18) %></td>
            <td width="20"><script>playring('<%=rsnews("id")%>')</script></td>
            <% 
if ii mod 2=0 then response.Write("</tr>") 
rsnews.MoveNext 
next
%>
          </tr>
        </table></td>
      </tr>
    </table>
        <%end if%>
        <img src="images/iphoto.gif" alt="" width="414" height="29" border="0" usemap="#MapMapMap2">
        <map name="MapMapMap2">
          <area shape="rect" coords="48,4,129,26" href="photo.asp">
        </map></td>
  </tr>
</table>
<table width="427"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="92" valign="top"><table width="100%" height="27%"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="422" height="68" valign="top"><table border="0" align="center" cellpadding="0" cellspacing="0">
          <%
cid=request("id")
if cid="" then
sql="select * from HN_pro order by id desc"
else
sql="select * from HN_pro where cstr(class)='"&cid&"' order by id desc"
end if
set rsnews=server.createobject("adodb.recordset")
rsnews.open sql,conn,1,1
	if rsnews.eof then 
	Response.Write "</td></tr></table>"
	end if
IF Not rsnews.eof Then
proCount=rsnews.recordcount
	rsnews.PageSize=24				  '定义显示数目
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

k=1
do while Not rsnews.eof and  k<2
%>
          <tr>
            <%for n=1 to 3%>
            <td><table width="126" height="68"  border="0" cellpadding="0" cellspacing="0" align="center">
              <tr>
                <td width="126"><a href="<%=rsnews("come")%>" rel="lightbox[plants]"><img src="<%=rsnews("come")%>" width="98" height="98" border="0"></a></td>
              </tr>
            </table></td>
            <%
rsnews.MoveNext 
if rsnews.eof then exit for
if rsnews.eof then exit do
next
%>
          </tr>
          <%
k=k+1
Loop
%>
        </table>
          <%end if%></td>
      </tr>
    </table>
        </td>
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