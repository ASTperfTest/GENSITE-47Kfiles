<!--#Include file="../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
 apath=server.mappath(HTUploadPath)
' response.write apath & "<HR>"
 Set upl = Server.CreateObject("SoftArtisans.FileUp")
 upl.Path = apath


if upl.form("submitTask") = "ADD" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		showErrBox()
	else
		doUpdateDB()
		showDoneBox()
	end if

else

	showHTMLHead()
	formFunction = "add"
	showForm()
	initForm()
	showHTMLTail()
end if
%>

<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm	'---- 新增時的表單預設值放在這裡
