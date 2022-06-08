<%@ CodePage = 65001 %>
<%
apath=server.mappath(session("uploadPath")) & "\"
Set xup = Server.CreateObject("TABS.Upload")
xup.codepage=65001
xup.Start apath

function xUpForm(xvar)
	xUpForm = xup.form(xvar)
end function



ACT = xUpForm("send")
	xPath = xUpForm("xxPath")
	upPath = session("uploadPath") & xPath
	if right(upPath,1)<>"/" then	upPath = upPath & "/"

If ACT = "新增" Then 

'Set fldr = fso.GetFolder(server.MapPath(upPath))
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.CreateFolder(server.MapPath(upPath & xUpForm("newFolder")))
%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache">
	<head>
	</head>
	<body>
<script language=vbs>
	msgBox "目錄 <%=server.MapPath(upPath & xUpForm("newFolder"))%> 建立成功"
	window.parent.navigate "index.asp"
'	window.navigate "fileMan.asp?xPath=<%=xPath%>"
</script>	
</body>
</html>
<%	
	Response.End
	
elseif ACT <> "送出" Then 
	Response.End
end if
	


'============  Photo Upload ==================

apath = server.MapPath(upPath)
	if right(apath,1)<>"/" then	apath = apath & "/"


For Each Form In xup.Form
     If Form.IsFile Then
        'Response.Write Form.Name & "=" & Form.FileName & "," & Form.SaveName
        xup.Form(Form.Name).Save(apath)
    End If

Next





%>
<html>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache">
	<head>
	</head>
	<body>
<script language=vbs>
	msgBox "檔案上傳至 --  <%=xPath%> -- 成功!"
	window.navigate "fileMan.asp?xPath=<%=xPath%>"
</script>
</body>
</html>