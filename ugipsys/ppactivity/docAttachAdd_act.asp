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
 
 pid=xUpForm("pid")
 filename_note=xUpForm("filename1_note")
 
  For each xatt in xup.Attachments
	  	ofname = xatt.FileName
  	  	    xatt.SaveFile apath & ofname, false
  Next	
  
  if ofname ="" then
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title></title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	doneStr="資料沒有填寫完整！"
	alert(doneStr)
	history.back()
</script>
</body>
</html>
  
<%
  response.end
  end if  
  sql="insert into docdownload_attach(pid,filename,filename_note) values('" & pid & "','" & ofname & "','" & filename_note & "')"
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
	doneStr="新增完成！"
	alert(doneStr)
	document.location.href = "docAttach.asp?id=<% =pid %>"
</script>
</body>
</html>


