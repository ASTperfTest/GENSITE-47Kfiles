<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#Include file = "../inc/index.inc" -->
<!--#Include file = "../inc/server.inc" -->
<!--#Include file = "../inc/dbFunc.inc" -->
<html><head>
<meta http-equiv="pragma: nocache">
<%
	SQLCom = "select * from CatTreeRoot Where CtRootID = "& ItemID 
	set RS = conn.execute(SqlCom)


	FtypeName = RS("CtRootName")


   const cid     = 2	'CatID
   const cname   = 3	'CatName
   const cparent = 7	'DataParent
   const cChild  = 8	'ChildCount
   const clevel  = 6	'DataLevel


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
<%
 If ForumType = "A" or ForumType = "B" Then %>
<script language=javascript>
<%	Response.write "treeRoot = gFld(""ForumToc"", """ & FtypeName & """, ""CatList.asp?ItemID="& ItemID &"&CatID=0"")" & vbcrlf
	NowCount = 0

	SqlCom = "SELECT * FROM CatTreeNode WHERE CtRootID = "& PkStr(ItemID,"") &" Order by DataLevel, CatShowOrder"
	set RS = conn.execute(SqlCom)
	while not RS.eof
		xParent = "treeRoot"
		if rs(clevel) > 1 then	xParent = "N" & RS(cparent)
		if RS("CtNodeKind") = "C" then
			CatLink = "CatList.asp?ItemID="& ItemID &"&CatID="& rs("CtNodeID")
			Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
		else
	     ForumLink = "ctXMLin.asp?ItemID="& ItemID &"&CtNodeID="& rs("CtNodeID")
	     Response.Write "insDoc(N"& rs(cparent) &", gLnk(""ForumToc"", """& rs("CatName") &""", """& ForumLink &"""))" & vbcrlf
		end if
		RS.moveNext
	wend

	SQLCom = "select * from CatTreeNode Where CtNodeKind='C' AND CtRootID = '"& ItemID &"' Order by DataLevel, CatShowOrder"
	Set RS = Conn.execute(SQLCom)
	If not rs.EOF Then %>
<%	 Do while not rs.EOF
	 NowCount = NowCount + 1
	 If ForumType = "A" Then
	  CatLink = "CatList.asp?ItemID="& ItemID &"&CatID="& rs("CtNodeID")
	  If rs(clevel) = 1 Then
'	   Response.Write "N"& rs(cid) &"= insFld(treeRoot, gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
	  Else
'	   Response.Write "N"& rs(cid) &"= insFld(N"& rs(cparent) &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
	  End IF
	 ElseIF ForumType = "B" Then
	   CatLink = "CatView.asp?ItemID="& ItemID &"&CatID="& rs("CtNodeID")
'	   Response.Write "N"& rs(cid) &"= insFld(treeRoot, gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
	 End IF
	  SQLForum = "Select * From CatTreeNode Where CtNodeKind='U' AND CtRootID =N'" & ItemID & "' AND DataParent ="& rs(cid) &" Order By CatShowOrder"
'	  response.write "//" & sqlForum & vbCRLF
	  Set RSForum = Conn.execute(SQLForum)
	   If Not rsforum.EOF Then
	    Do while not rsforum.EOF
	     ForumLink = "ForumView.asp?ItemID="& ItemID &"&ForumID="& rsforum("CtNodeID")
'	     Response.Write "insDoc(N"& rs(cid) &", gLnk(""ForumToc"", """& rsforum("CatName") &""", """& ForumLink &"""))" & vbcrlf
	    rsforum.movenext
     	Loop
	   End If
	 RS.movenext
     Loop
	End If 

	
	
%>
     initializeDocument();
</script>
<%
 ElseIf ForumType = "C" Then %>
<script language=javascript>
<% 	  Response.write "treeRoot = gFld(""ForumToc"", """ & FtypeName & """, ""CatView.asp?ItemID="& ItemID &"&CatID=0"")" & vbcrlf
	  SQLForum = "Select * From CatTreeNode Where ItemID =N'"& ItemID &"' Order By ForumShowOrder"
	  Set RSForum = Conn.execute(SQLForum)
	   If Not rsforum.EOF Then %>
<%	    Do while not rsforum.EOF
	     ForumLink = "ForumView.asp?ItemID="& ItemID &"&ForumID="& rsforum("ForumID")
	     Response.Write "insDoc(treeRoot, gLnk(""ForumToc"", """& rsforum("ForumName") &""", """& ForumLink &"""))" & vbcrlf
	    rsforum.movenext
     	Loop
	   End If %>
     initializeDocument();
</script>
<%
 ElseIF ForumType = "D" Then
 	  SQLForum = "Select * From Forum Where ItemID =N'"& ItemID &"'"
	  Set RSForum = Conn.execute(SQLForum) %>
<script language=javascript>
	window.parent.ForumToc.location = "ForumView.asp?ForumID=<%=rsforum("ForumID")%>&ItemID=<%=ItemID%>"
</script>
<%
 End IF %>
</body></html>
