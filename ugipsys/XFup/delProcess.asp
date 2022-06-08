<%@ CodePage = 65001 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
ACT = request("submitTask")
	xPath =request("xxPath")
	upPath = session("uploadPath") & xPath
	if right(upPath,1)<>"/" then	upPath = upPath & "/"
'	response.write upPath & "<HR>"
'	response.write ACT & (Act= "刪檔") & "<HR>"

if ACT = "delFiles" Then 
'	response.write request("insertFiles")
'	response.write "<HR>"
	
	Set fso = CreateObject("Scripting.FileSystemObject")

	fList = split(request("insertFiles"),", ")
	for x = 0 to UBound(fList,1)
'		response.write fList(x) & "<HR>"
		fso.DeleteFile(server.MapPath(upPath & fList(x)))
	next
%>
<script language=vbs>
	msgBox "刪除 <%=UBound(fList,1)+1%> 個檔案成功"
	window.navigate "fileMan.asp?xPath=<%=xPath%>"
</script>	
<%	
	response.end
elseif ACT = "delDir" Then 
	if right(upPath,1)<>"/" then	upPath = upPath & "/"
	upPath = left(upPath, len(upPath)-1)
	response.write upPath & "<HR>"
	xPos = inStrRev(upPath, "/")
	if xPos < 1 then	response.end
	
	parentPath = left(upPath, xPos-1)
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.DeleteFolder(server.MapPath(upPath))	
%>
<script language=vbs>
	msgBox "刪除目錄成功"
	window.parent.navigate "index.asp"
'	window.navigate "fileMan.asp?xPath=<%=parentPath%>"
</script>	
<%	
	
'	response.write upPath & "<HR>"
	Response.End
end if


%>
