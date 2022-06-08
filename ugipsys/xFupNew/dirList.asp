<%@ CodePage = 65001 %><% 
    Response.Expires = 0
    Response.Charset = "utf-8"
    HTProgCode = "GC1AP5" 
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=response.charset%>">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache">

<script language=javascript>
 var OpenFolder = "/imageFolder/ftv2folderclosed.gif"
 var CloseFolder = "/imageFolder/ftv2folderopen.gif"
 var DocImage = "/imageFolder/openfold.gif"
</script>
<script src="/imageFolder/FolderTree.js"></script>
<base target="Catalogue">
<title>dir list</title>
<style>
.FolderStyle { text-decoration: none; color:#000000}
</style>
</head>
<body leftmargin="3" topmargin="3">
<script language=javascript>
<%
dim NowCount
Set fso = server.CreateObject("Scripting.FileSystemObject")


xPath = request("xPath")
if xPath = "" then 
    xPath = "/"
end if
if right(xPath,1)<>"/" then	
    xPath = xPath & "/"
end if


Response.write "treeRoot = gFld(""ForumToc"", """ & session("uploadPath") & """, ""fileMan.asp?xpath="")" & vbcrlf
	NowCount = 0

	traverseFolder server.MapPath(xpath), "treeRoot", ""


%>
     initializeDocument();
</script>
</body></html>


<%
sub traverseFolder(zPath,zParent,vPath)
	Set fldr = fso.GetFolder(zPath)

	  np = vPath
		if right(np,1)<>"/" then	np = np & "/"
		if left(np,1)="/" then	np = mid(np,2)
	for each sf in fldr.SubFolders
	  CatLink = "fileList.aspx?path="& server.URLEncode(sf.path)
	  CatLink = "fileMan.asp?xpath="& np & server.URLEncode(sf.name)
'	  CatLink = "fileList.asp?path="& server.URLEncode(sf.path)
	  set grandson = sf.SubFolders
	  if grandson.count > 0 then
		NowCount = NowCount + 1
		Response.Write "N"& NowCount &"= insFld(" & zParent &", gFld(""ForumToc"", """& sf.name & """, """& CatLink &"""))" & vbcrlf
			traverseFolder sf.path , "N"& NowCount, np & sf.name
	  else
		Response.Write "insDoc("& zParent &", gLnk(""ForumToc"", """& sf.name  &""", """& CatLink &"""))" & vbcrlf
	  end if
	next	
end sub
%>