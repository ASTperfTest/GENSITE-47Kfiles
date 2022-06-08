<%@ CodePage = 65001 %>
<%
'Randomize
Set xup = Server.CreateObject("LyfUpload.UploadFile")
ACT = xup.request("send")
	xPath = xup.request("xxPath")
	upPath = session("uploadPath") & xPath
	if right(upPath,1)<>"/" then	upPath = upPath & "/"

If ACT = "新增" Then 
'	response.write xup.request("newFolder")

	response.write upPath & xup.request("newFolder")
'response.end
'Set fldr = fso.GetFolder(server.MapPath(upPath))
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.CreateFolder(server.MapPath(upPath & xup.request("newFolder")))
%>
<script language=vbs>
	msgBox "目錄 <%=server.MapPath(upPath & xup.request("newFolder"))%> 建立成功"
	window.parent.navigate "index.asp"
'	window.navigate "fileMan.asp?xPath=<%=xPath%>"
</script>	
<%	
	Response.End
	
elseif ACT <> "送出" Then 
	Response.End
end if
	


'============  Photo Upload ==================

apath = server.MapPath(upPath)
fn = xup.SaveFile ("photo1", apath, true, filename)
fn = xup.SaveFile ("photo2", apath, true, filename)
fn = xup.SaveFile ("photo3", apath, true, filename)
fn = xup.SaveFile ("photo4", apath, true, filename)
fn = xup.SaveFile ("photo5", apath, true, filename)
fn = xup.SaveFile ("photo6", apath, true, filename)
fn = xup.SaveFile ("photo7", apath, true, filename)
fn = xup.SaveFile ("photo8", apath, true, filename)
fn = xup.SaveFile ("photo9", apath, true, filename)
fn = xup.SaveFile ("photo0", apath, true, filename)


%>
<script language=vbs>
	msgBox "檔案上傳至 --  <%=xPath%> -- 成功!"
	window.navigate "fileMan.asp?xPath=<%=xPath%>"
</script>
