<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "Pn02M02" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<html><head>
<link href="/css/tree.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<meta http-equiv="pragma: nocache">
<%

%>
<script language=javascript>
 var OpenFolder = "/imageFolder/ftv2folderclosed.gif"
 var CloseFolder = "/imageFolder/ftv2folderopen.gif"
 var DocImage = "/imageFolder/jobTitle.gif"
</script>
<script src="/imageFolder/FolderTree.js"></script>
<base target="Catalogue">
<title></title>
<style>
.FolderStyle { text-decoration: none; color:#000000}
</style>
</head><body leftmargin="3" topmargin="3">
<script language=javascript>
<%	
	SqlCom = "SELECT * FROM Dept Order by len(deptID), seq"
	set RS = conn.execute(SqlCom)

	Response.write "treeRoot = gFld(""ForumToc"", """ & RS("deptName") & """, ""deptEdit.asp?deptID="& RS("deptID") & """)" & vbcrlf
	RS.moveNext
	NowCount = 0

	while not RS.eof
		xParent = "treeRoot"
		if len(rs("deptID")) > 2 then	xParent = "N" & RS("parent")
		if RS("nodeKind") = "D" then
			CatLink = "deptEdit.asp?deptID="& RS("deptID") 
			Response.Write "N"& RS("deptID") &"= insFld(" & xParent &", gFld(""ForumToc"", """& RS("deptName") &""", """& CatLink &"""))" & vbcrlf
		else
	     ForumLink = "jobEdit.asp?deptID="& RS("deptID")
	     Response.Write "insDoc("& xParent &", gLnk(""ForumToc"", """& rs("deptName") &""", """& ForumLink &"""))" & vbcrlf
		end if
		RS.moveNext
	wend
	
%>
     initializeDocument();
</script>
</body></html>
