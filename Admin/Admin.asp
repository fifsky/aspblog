<!--#include file="conn.asp"-->
<!--#include file="error.asp"-->
<% dim rs,sql
	set rs=server.createobject("adodb.recordset") 
	sql="select top 1 * from admin"
 	rs.open sql,conn,1,3  
%>
<%
	'����
	Sub Header()
		Response.Write"<link href=images/admin.css rel=stylesheet>"
		Response.Write"<title>��վ��̨����</title>"
		Response.Write"<META http-equiv=Content-Type content=text/html;charset=GB2312>"
		%>
		<Script language="javascript">
		//����ȷ��
		function checkclick(msg){
		if(confirm(msg))
			{
			event.returnValue=true;}else{event.returnValue=false;
			}
		}
		</Script>
		<%
	End Sub
	Sub Footer()
		Response.Write"<br><br>"
		Response.Write"<hr size=0 noshade width=80% color=#698CC3>"
		Response.Write"<center><font style='font-size: 11px; font-family: Tahoma, Verdana, Arial'>COPYRIGHT 2006~2008 <a href=http://www.vtsj.cn  target=_blank>VTsj.CN</a> </font></center>"
		Response.Write"</body>"
		Response.Write"</html>"
		Response.End
		Set Rs=Nothing
		Set Rs1=Nothing
		Set Conn=Nothing
	End Sub

Dim Admin_Title
Admin_Title="��̨����"
Header()
Dim Admin_Class
Admin_Class=",1,"
Select Case Request("menu")
	Case "leftbody"
		leftbody
	Case "pass"
		pass
	Case "topbanner"
		topbanner
	Case "out"
		session("admin_name")=""
		response.redirect "Index.asp"
	Case Else
		index
End Select

Sub topbanner
	Dim MSCode
	If IsSqlDataBase = 1 Then
		MSCode="SQL"
	Else
		MSCode="ACC"
	End If
%>
<body topmargin="0" rightmargin="0" leftmargin="0">
<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="100%" class=a2>
	<tr class="a1">
		<td width="63%" align="center">
		<b>��Ȩ���У�VTSJ.CN</b>  |  <a href=?menu=pass  target="main">�ص�������ҳ</a></td>
		<td width="37%"><a href=../ target="_blank">վ����ҳ</a></td>
	</tr>
</table>
<%
End Sub


Sub pass

%>
<body Style="background-color:#9EB6D8" text="#000000" leftmargin="5" topmargin="5">
<br>
<br>

<table cellpadding="3" cellspacing="1" border="0" width="90%" class="a2">
  <tr>
    <td class=a1>��������Ϣ</td>
  </tr>
  <tr>
    <td class=a4><table cellpadding="3" cellspacing="0" border="0" width="100%">
        <tr class=a4>
          <td width="20%">����������</td>
                        <td width="29%"><%=Request.ServerVariables("SERVER_NAME")%></td>
                        <td width="21%">������IP��</td>
                        <td width="30%"><%=Request.ServerVariables("LOCAL_ADDR")%></td>
        </tr>
                      <tr class=a3>
                        <td>�������˿ڣ�</td>
                        <td><%=Request.ServerVariables("SERVER_PORT")%></td>
                        <td>������ʱ�䣺</td>
                        <td><%=now%></td>
        </tr>
                      <tr class=a4>
                        <td>���ļ�����·����</td>
                        <td colspan="3"><%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
        </tr>
                      <tr class=a3>
                        <td>����������ϵͳ��</td>
                        <td><%=Request.ServerVariables("OS")%></td>
                        <td>������CPU������</td>
                        <td><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%></td>
        </tr>
                      <tr class=a4>
                        <td>�ͻ���IP��</td>
                        <td><%=Request.ServerVariables("REMOTE_ADDR")%></td>
                        <td>�˿ڣ�</td>
                        <td><%=Request.ServerVariables("REMOTE_PORT")%></td>
        </tr>
                      <tr class=a3>
                        <td>����</td>
                        <td><%=Request.ServerVariables("HTTP_X_FORWARDED_FOR")%></td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
        </tr>
                      <tr class=a4>
                        <td>�ű��������棺</td>
                        <td><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
                        <td>Jmail�������֧�֣�</td>
                        <td><%If IsObjInstalled("JMail.Message") Then%>
      Jmail4.3�������֧��
        <%elseIf IsObjInstalled("JMail.SMTPMail") then%>
      Jmail4.2���֧��
      <%else%>
      ��֧��Jmail���
      <%end if%></td>
        </tr>
                      <tr class=a3>
                        <td>FSO���֧�֣�</td>
                        <td><%If IsObjInstalled("Scripting.FileSystemObject") then%>
      FSO֧��
        <%else%>
      ��֧��FSO���
      <%end if%>
 </td>
                        <td>CDONTS�������֧�֣�</td>
                        <td><%If IsObjInstalled("CDONTS.NewMail") then%>
      CDONTS�������֧��
        <%else%>
      ��֧��CDONTS�������
      <%end if%></td>
        </tr>
    </table></td>
  </tr>
</table>

  <%
	Footer()
end sub

sub leftbody
%>
    <script language="JavaScript">
function ClearAllDeploy(){
	var deployitem=FetchCookie("deploy");
	var admin_start;
	var userdeploy='';
	admin_start= deployitem ? deployitem.indexOf("\n") : -1;
	if(admin_start!=-1){
		userdeploy=deployitem.substring(0,admin_start);
	}
	for(i=0;i<20;i++){
		obj=document.getElementById("cate_"+"id"+i);	
		img=document.getElementById("img_"+"id"+i);
		if(obj && obj.style.display=="none"){
			obj.style.display="";
			img_re=new RegExp("_open\\.gif$");
			img.src=img.src.replace(img_re,'_fold.gif');
		}
	}
	deployitem=userdeploy+"\n\t\t";
	SetCookie("deploy",deployitem);
}
function SetAllDeploy(){
	var deployitem=FetchCookie("deploy");
	var admin_start;
	var userdeploy='';
	var admindeploy='';
	var i;
	admin_start= deployitem ? deployitem.indexOf("\n") : -1;
	if(admin_start!=-1){
		userdeploy=deployitem.substring(0,admin_start);
	}
	for(i=0;i<20;i++){
		obj=document.getElementById("cate_"+"id"+i);	
		img=document.getElementById("img_"+"id"+i);
		if(obj && obj.style.display==""){
			obj.style.display="none";
			img_re=new RegExp("_fold\\.gif$");
			img.src=img.src.replace(img_re,'_open.gif');
		}
		admindeploy=admindeploy+"id"+i+"\t";
	}
	deployitem=userdeploy+"\n\t"+admindeploy;
	SetCookie("deploy",deployitem);
}
function IndexDeploy(ID,type){
	obj=document.getElementById("cate_"+ID);	
	img=document.getElementById("img_"+ID);
	if(obj.style.display=="none"){
		obj.style.display="";
		img_re=new RegExp("_open\\.gif$");
		img.src=img.src.replace(img_re,'_fold.gif');
		SaveDeploy(ID,type,false);
	}else{
		obj.style.display="none";
		img_re=new RegExp("_fold\\.gif$");
		img.src=img.src.replace(img_re,'_open.gif');
		SaveDeploy(ID,type,true);
	}
	return false;
}
function SaveDeploy(ID,type,is){
	var foo=new Array();
	var deployitem=FetchCookie("deploy");
	var admin_start;
	var admindeploy='';
	var userdeploy='';
	admin_start= deployitem ? deployitem.indexOf("\n") : -1;
	if(admin_start!=-1){
		admindeploy= deployitem.substring(admin_start+1,deployitem.length);
		userdeploy = deployitem.substring(0,admin_start);
	}
	if(deployitem!=null){
		if(admin_start!=-1){
			deployitem = type==0 ? userdeploy : admindeploy;
		}
		deployitem=deployitem.split("\t");
		for(i in deployitem){
			if(deployitem[i]!=ID && deployitem[i]!=""){
				foo[foo.length]=deployitem[i];
			}
		}
	}
	if(is){
		foo[foo.length]=ID;
	}
	deployitem = type==0 ? "\t"+foo.join("\t")+"\t\n"+admindeploy : userdeploy+"\n\t"+foo.join("\t")+"\t";
	SetCookie("deploy",deployitem)
}
function SetCookie(name,value){
	expires=new Date();
	expires.setTime(expires.getTime()+(86400*365));
	document.cookie=name+"="+escape(value)+"; expires="+expires.toGMTString()+"; path=/";
}
function FetchCookie(name){
	var start=document.cookie.indexOf(name);
	var end=document.cookie.indexOf(";",start);
	return start==-1 ? null : unescape(document.cookie.substring(start+name.length+1,(end>start ? end : document.cookie.length)));
}
    </script>
  <body Style="background-color:#9EB6D8" text="#000000" leftmargin="5" topmargin="5">
</p>
<table width="100%" cellspacing="1" cellpadding="4" border="0" class=a4>
  <tr>
    <td class="a3"><table width="98%" cellspacing="1" cellpadding="4" class="a2">
        <tr>
          <td class="a4"><a href="#" onClick="return ClearAllDeploy()" class="a_bold">[չ��+]</a>&nbsp;&nbsp;<a href="#" onClick="return SetAllDeploy()" class="a_bold">[�ر�-]</a> </td>
        </tr>
      </table></td>
  </tr>
	<tr><td>
		<table width="98%" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
		<a style="float:right"><img src="images/cate_fold.gif" border=0></a>
			<a href="?menu=pass" class="a1" target="main"><b>������ҳ</b></a>
		</td></tr>
		</table>
	</td></tr>
	<tr><td>
		<table width="98%" align="center" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onClick="return IndexDeploy('id1',1)"><img id="img_id1" src="images/cate_fold.gif" border=0></a>
			<a href="#" onClick="return IndexDeploy('id1',1)" class="a1"><b>ϵͳ����</b></a>
		</td></tr>
		<tbody id="cate_id1" style="display:none;">	
		<tr>
		  <td class="a4" align="center">
		  <a href="admin_web.asp" target="main">��վ��������</a>
		  </td>
		</tr>
		<tr>
		  <td class="a3" align="center">
		  <a href="admin_passwd.asp" target="main">����Ա����</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4" align="center">
		  <a href="UploadFileManage.asp" target="main">�ϴ��ļ�����</a>
		  </td>
		</tr>
		</tbody>
		</table>
	</td></tr>
	<tr><td>
		<table width="98%" align="center" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onClick="return IndexDeploy('id2',1)"><img id="img_id2" src="images/cate_fold.gif" border=0></a>
			<a href="#" onClick="return IndexDeploy('id2',1)" class="a1"><b>���ݹ���</b></a>
		</td></tr>
		<tbody id="cate_id2" style="display:">
		<tr>
		  <td class="a4" align="center">
		  <a href="admin_news.asp?work=up&types=3" target="main">�ҵ�����</a>
		  </td>
		</tr>
		<tr>
		  <td class="a3" align="center">
		  <a href="admin_pro.asp?work=palycla" target="main">������</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4" align="center">
		  <a href="admin_new.asp?work=palycla" target="main">�ռǹ���</a>
		  </td>
		</tr>
		<tr>
		  <td class="a3" align="center">
		  <a href="admin_down.asp?work=palycla" target="main">��������</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4" align="center">
		  <a href="GuestBook.asp" target="main">���Թ���</a>
		  </td>
		</tr>
		</tbody>
		</table>
	</td></tr>

	<tr><td>
		<table width="98%" align="center" cellspacing="1" cellpadding="4" class="a2">
		<tr><td class="a1">
			<a style="float:right" href="#" onClick="return IndexDeploy('id8',1)"><img id="img_id8" src="images/cate_fold.gif" border=0></a>
			<a href="#" onClick="return IndexDeploy('id8',1)" class="a1"><b>��������</b></a>
		</td></tr>
		<tbody id="cate_id8" style="display:">
		<tr>
		  <td class="a4" align="center">
		  <a href="../" target="_blank">��վ��ҳ</a>
		  </td>
		</tr>
		<tr>
		  <td class="a3" align="center">
		  <a target="_top" href="?menu=out">�˳�����</a>
		  </td>
		</tr>
		<tr>
		  <td class="a4" align="center"><a target="_blank" href="http://www.vtsj.cn/">�ٷ���վ</a>		  </td>
		</tr>

		</tbody>
		</table>
	</td></tr>
	<tr>
    <td height="2"></td>
  </tr>

</table>
<%
end sub


Sub index
%>
<frameset cols="160,*" frameborder="no" border="0" framespacing="0" rows="*">
<frame name="menu" noresize scrolling="yes" src="?menu=leftbody">
<frameset rows="20,*" frameborder="no" border="0" framespacing="0" cols="*">
<frame name="a1" noresize scrolling="no" src="?menu=topbanner">
<frame name="main" noresize scrolling="yes" src="?menu=pass">
</html>
<%
end sub
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>