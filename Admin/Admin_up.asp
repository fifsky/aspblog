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
set upload=new upload_5xSoft ''建立上传对象
						
formPath="../upImgFile/" 

	''列出所有上传了的文件
    for each formName in upload.file
      set file=upload.file(formName)
	  if file.filesize>0 then
      if file.filesize>10000000 then
		response.write "<font size=2>图片大小超小了限制[<a href=# onclick=history.go(-1)>重新上传</a>]</font>"
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
         response.write "<font size=2>文件格式限制[<a href=# onclick=history.go(-1)>请重新上传</a>]</font>"
         response.end
      end if
	end if
						
	filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
	
    if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
       file.SaveAs Server.mappath(formpath&filename)   ''保存文件
       'response.write file.FilePath&file.FileName&"("&file.FileSize&") => "&formPath&File.FileName&"上传成功<br>"
       response.write "<font size=2>上传成功 <a href=# onclick=history.go(-1)>请返回</a>" 

	   end if
set file=nothing
next
set upload=nothing
Response.Write "<script>parent.add.come.value='upImgFile/"&FileName&"'</script>"
%>
