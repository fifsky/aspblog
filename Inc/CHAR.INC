<%
function htmlencode(str)
    dim result
    dim l
    if isNULL(str) then 
       htmlencode=""
       exit function
    end if
    l=len(str)
    result=""
	dim i
	for i = 1 to l
	    select case mid(str,i,1)
	           case "<"
	                result=result+"&lt;"
	           case ">"
	                result=result+"&gt;"
              case chr(13)
	                result=result+"<br>"
	           case chr(34)
	                result=result+"&quot;"
	           case "&"
	                result=result+"&amp;"
              case chr(32)	           
	                'result=result+"&nbsp;"
	                if i+1<=l and i-1>0 then
	                   if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then	                      
	                      result=result+"&nbsp;"
	                   else
	                      result=result+" "
	                   end if
	                else
	                   result=result+"&nbsp;"	                    
	                end if
	           case chr(9)
	                result=result+"    "
	           case else
	                result=result+mid(str,i,1)
         end select
       next 
       htmlencode=result
   end function


function sustainhtml(str)
    dim result
    dim l
    if isNULL(str) then 
       sustainhtml=""
       exit function
    end if
    l=len(str)
    result=""
	dim i
	for i = 1 to l
	    select case mid(str,i,1)
	           case chr(13)
	                result=result+"<br>"
	           case chr(34)
	                result=result+"&quot;"
                   case chr(32)	           
	                'result=result+"&nbsp;"
	                if i+1<=l and i-1>0 then
	                   if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then	                      
	                      result=result+"&nbsp;"
	                   else
	                      result=result+" "
	                   end if
	                else
	                   result=result+"&nbsp;"	                    
	                end if
	           case "&"
	                result=result+"&amp;"
	           case chr(9)
	                result=result+"    "
	           case else
	                result=result+mid(str,i,1)
         end select
       next 
       sustainhtml=result
   end function

 ' 检查sql字符串中是否有单引号，有则进行转化
   function CheckStr(str)
       dim tstr,l,i,ch
	   l=len(str)
	   for i=1 to l
	       ch=mid(str,i,1)
	       if ch="'" then
		      tstr=tstr+"'"
		   end if
		   tstr=tstr+ch
	   next
	   CheckStr=tstr
   end function

'email有效性检查
function IsValidEmail(email)

dim names, name, i, c

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
end if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
end if
if InStr(email, "..") > 0 then
   IsValidEmail = false
end if

end function

'将一文字中的HTML的标识去掉
Function FilterHTML(strToFilter)
Dim strTemp
strTemp = strToFilter
While Instr(1,strTemp,"<") AND Instr(1, strTemp, ">")
strTemp = Left(strTemp, Instr(1, strTemp, "<")-1) & Right(strTemp, Len(strTemp)-Instr(1,strTemp, ">"))
WEnd
FilterHTML = strTemp
End Function

%>