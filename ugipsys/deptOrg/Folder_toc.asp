<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "Pn02M02" 
' ============= Modified by Chris, 2006/08/22, to handle "DeptIdDouble" ========================
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/tree.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<meta http-equiv="pragma: nocache">
<script language=javascript>
 var OpenFolder = "/imageFolder/ftv2folderclosed.gif"
 var CloseFolder = "/imageFolder/ftv2folderopen.gif"
 var DocImage = "/imageFolder/openfold.gif"
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
' ===begin========== Modified by Chris, 2006/08/22, to handle "DeptIdDouble" ========================
	dblDeptIDlen=2
 	if checkGIPconfig("DeptIdDouble") then	dblDeptIDlen=2

	SqlCom = "SELECT * FROM Dept where (deptId = Left('" & session("deptID") & "',Len(deptId)) or deptId LIKE '" & session("deptID") & "%') Order by len(deptId), seq"
	set RS = conn.execute(SqlCom)

	Response.write "treeRoot = gFld(""ForumToc"", """ & RS("deptName") & """, ""deptEdit.asp?phase=edit&deptID="& RS("deptID") & """)" & vbcrlf
	RS.moveNext
	NowCount = 0

	while not RS.eof
		xParent = "treeRoot"
		if len(rs("deptID")) > (dblDeptIDlen+1) then	xParent = "N" & RS("parent")
' ===end========== Modified by Chris, 2006/08/22, to handle "DeptIdDouble" ========================
		if RS("nodeKind") = "D" then
			CatLink = "deptEdit.asp?deptID="& RS("deptID") &"&phase=edit"
			Response.Write "N"& RS("deptID") &"= insFld(" & xParent &", gFld(""ForumToc"", """& RS("deptName") &""", """& CatLink &"""))" & vbcrlf
		else
			CatLink = "deptEdit.asp?deptID="& RS("deptID") &"&phase=edit"
			Response.Write "N"& RS("deptID") &"= insFld(" & xParent &", gFld(""ForumToc"", """& RS("deptName") &""", """& CatLink &"""))" & vbcrlf
'	     ForumLink = "jobEdit.asp?deptID="& RS("deptID")
'	     Response.Write "insDoc("& xParent &", gLnk(""ForumToc"", """& rs("deptName") &""", """& ForumLink &"""))" & vbcrlf
		end if
		RS.moveNext
	wend
	
%>
     initializeDocument();
</script>
</body></html>
