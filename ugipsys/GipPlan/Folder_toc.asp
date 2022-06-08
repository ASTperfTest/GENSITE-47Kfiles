<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
<link href="/css/tree.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<meta http-equiv="pragma: nocache">
<%
	SQLCom = "select * from CatTreeRoot Where ctRootId = "& ItemID 
	set RS = conn.execute(SqlCom)


	FtypeName = RS("CtRootName")


   const cid     = 2	'CatID
   const cname   = 3	'CatName
   const cparent = 7	'DataParent
   const cChild  = 8	'ChildCount
   const clevel  = 6	'DataLevel
'	SqlCom = "SELECT * FROM CatTreeNode WHERE ctRootId = "& PkStr(ItemID,"") &" Order by DataLevel, CatShowOrder"
'	response.write sqlCom

%>
<script language=javascript>
 var OpenFolder = "image/ftv2folderclosed.gif"
 var CloseFolder = "image/ftv2folderopen.gif"
 var DocImage = "image/openfold.gif"
</script>
<script src="FolderTree.js"></script>
<base target="Catalogue">
<title></title>
<style>
.FolderStyle { text-decoration: none; color:#000000}
</style>
</head><body leftmargin="3" topmargin="3">
<%  if Instr(session("uGrpID")&",", "HTSD,") > 0 then %> 
<A href="genXML.asp?itemID=<%=itemID%>" target="ForumToc">genXML</A><BR/>
<%	end if %>
<script language=javascript>
<%	Response.write "treeRoot = gFld(""ForumToc"", """ & FtypeName & """, ""CatList.asp?ItemID="& ItemID &"&CatID=0"")" & vbcrlf
	NowCount = 0

	SqlCom = "SELECT * FROM CatTreeNode WHERE ctRootId = "& PkStr(ItemID,"") &" Order by dataLevel, catShowOrder"
	set RS = conn.execute(SqlCom)
	while not RS.eof
		xParent = "treeRoot"
		if rs(clevel) > 1 then	xParent = "N" & RS(cparent)
		if RS("CtNodeKind") = "C" then
			CatLink = "CatList.asp?ItemID="& ItemID &"&CatID="& rs("CtNodeID")
			Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
		else
	     ForumLink = "CtUNitNodeEdit.asp?phase=edit&CtNodeID=" & rs("CtNodeID")
	     Response.Write "insDoc("& xParent &", gLnk(""ForumToc"", """& rs("CatName") &""", """& ForumLink &"""))" & vbcrlf
		end if
		RS.moveNext
	wend
	
%>
     initializeDocument();
</script>
</body></html>
