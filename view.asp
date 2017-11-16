<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="include.asp"-->
<%
dim dbid,sql
dbid=FormatSQL(SafeRequest("id",1))
set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from HN_news where id="&dbid
rs.open sql,conn,1,3
rs("hotclick")=rs("hotclick")+1
rs.update
demo=replace(rs("demo"),chr(10)+chr(13),"<br>")%> 

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=rs("title")%> - <%=WebName%></title>
<meta name="keywords" content="<%=W_WebSiteKeyword%>">
<LINK href="css.css" type=text/css rel=stylesheet>
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
	  <td width="427" height="414" colspan="5" rowspan="2" valign="top" background="images/VTsj_17.jpg"><table width="423"  border="0" cellpadding="5" cellspacing="0">

        <tr>
          <td width="413"><SCRIPT LANGUAGE="javascript">
function scroll(n)
{temp=n;
Out1.scrollTop=Out1.scrollTop+temp;
if (temp==0) return;
setTimeout("scroll(temp)",80);
}
      </SCRIPT>
              <TABLE WIDTH="87%">
                <TR>
                  <TD VALIGN="TOP" ><DIV ID=Out1 STYLE="width:405; height:370;overflow: hidden">
                      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td height="10"></td>
                        </tr>
                        <tr>
                          <td class="style3" style="line-height: 20px"><%=demo%></td>
                        </tr>
                      </table>
                  </DIV></TD>
                </TR>
                <TR>
                  <TD align="right" VALIGN="TOP" ><IMG SRC="images/up.gif" WIDTH="50" HEIGHT="20" onMouseOver="scroll(-1)" onMouseOut="scroll(0)" onMouseDown="scroll(-3)" BORDER="0" ALT="按下鼠标速度会更快！"> <IMG SRC="images/down.gif" onMouseOver="scroll(1)" onMouseOut="scroll(0)"  onmousedown="scroll(3)" BORDER="0" WIDTH="50" HEIGHT="20" ALT="按下鼠标速度会更快！"> </TD>
                </TR>
            </TABLE></td>
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