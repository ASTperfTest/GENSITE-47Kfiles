<% Response.Expires = 0
HTProgCap="課程"
HTProgFunc="編修"
HTProgCode="PA001"
HTUploadPath=session("Public")+"class/"
HTProgPrefix="paAct" %>
<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<% response.expires = 0 %>
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<!--#include virtual = "/inc/hyftdGIP.inc" -->
<%

apath=server.mappath(HTUploadPath) & "\"
set xup = Server.CreateObject("UpDownExpress.FileUpload")
xup.Open 
function xUpForm(xvar)
on error resume next
	xStr = ""
	arrVal = xup.MultiVal(xvar)
	for i = 1 to Ubound(arrVal)
		xStr = xStr & arrVal(i) & ", "
'		Response.Write arrVal(i) & "<br>" & Chr(13)
	next 
	if xStr = "" then
		xStr = xup(xvar)
		xUpForm = xStr
	else
		xUpForm = left(xStr, len(xStr)-2)
	end if
end function

id=xUpForm("id")
pid=xUpForm("pid")

if  xUpForm("submit")="修改" then

 filename1_note=xUpForm("filename1_note")
 ofilename1=xUpForm("ofilename1")
 
  For each xatt in xup.Attachments
	  	ofname = xatt.FileName
	  	if ofname <> ofilename1 and ofname<>"" then
	  	    if xup.IsFileExist( apath & ofilename1) then 
                      xup.DeleteFile apath & ofilename1
                    end if
  	  	    xatt.SaveFile apath & ofname, false
  	  	    sql="update docdownload_attach set filename='" & ofname & "' where id=" & id
                    conn.Execute(sql)
  	  	end if    
  Next	
  
  sql="update docdownload_attach set filename_note='" & filename1_note & "' where id=" & id
  conn.Execute(sql)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	doneStr="修改完成！"
	alert(doneStr)
	document.location.href = "docAttach.asp?id=<% =pid %>"
</script>
</body>
</html>

<% else 
 sql="delete docdownload_attach where id=" & id
  conn.Execute(sql)
  ofilename1=xUpForm("ofilename1")
 if xup.IsFileExist( apath & ofilename1) then 
 xup.DeleteFile apath & ofilename1
 end if 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	doneStr="刪除完成！"
	alert(doneStr)
	document.location.href = "docAttach.asp?id=<% =pid %>"
</script>
</body>
</html>  
<% end if %>
