<!--#include file="UPLOAD.INC"-->
<!--#include file="Error.asp"-->
<style type="text/css">
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>

<%
dim arr(3)
dim upload,file,formName,formPath,iCount,filename,fileExt,i
set upload=new upload_5xSoft ''�����ϴ�����
						
formPath="../upImgFile/" 

	''�г������ϴ��˵��ļ�
    for each formName in upload.file
      set file=upload.file(formName)
	  if file.filesize>0 then
      if file.filesize>10000000 then
		response.write "<font size=2>ͼƬ��С��С������[<a href=# onclick=history.go(-1)>�����ϴ�</a>]</font>"
		response.end
      end if



fileExt=lcase(right(file.filename,4))
 uploadsuc=false
Forum_upload="gif,jpg,jpeg,bmp,png,rar,zip,doc,rm,mp3,wav,mid,midi,ra,avi,mpg,mpeg,asf,asx,wma,mov"
 Forumupload=split(Forum_upload,",")
 for i=0 to ubound(Forumupload)
    if fileEXT="."&trim(Forumupload(i)) then
    uploadsuc=true
    exit for
    else
    uploadsuc=false
    end if
 next


if uploadsuc=false then
         response.write "<font size=2>�ļ���ʽ����[<a href=# onclick=history.go(-1)>�������ϴ�</a>]</font>"
         response.end
      end if
	end if
						
	filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
	
    if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
       file.SaveAs Server.mappath(formpath&filename)   ''�����ļ�
       'response.write file.FilePath&file.FileName&"("&file.FileSize&") => "&formPath&File.FileName&"�ϴ��ɹ�<br>"
       response.write "<font size=2>�ϴ��ɹ� <a href=# onclick=history.go(-1)>�뷵��</a>" 

	   end if
set file=nothing
next
set upload=nothing
Response.Write "<script>parent.add.come.value='upImgFile/"&FileName&"'</script>"
%>
