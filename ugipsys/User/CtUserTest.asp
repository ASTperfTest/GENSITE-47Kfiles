<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<html><head>
<meta http-equiv="pragma: nocache">
<%
	ItemID = 4		'-- 要處理哪個 tree
	
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
 var OpenFolder = "/imageFolder/ftv2folderclosed.gif"
 var CloseFolder = "/imageFolder/ftv2folderopen.gif"
 var DocImage = "/imageFolder/openfold.gif"
</script>
<script src="/imageFolder/FolderTree.js"></script>
<title></title>
<style>
.FolderStyle { text-decoration: none; color:#000000}
</style>
</head>
<body leftmargin="3" topmargin="3">

上稿權限設定<HR>

<%
  userID = request("userID")
%>
<form name="form1" method="post" action="CtUserSet_act.asp">
  <input type="hidden" name="userID" value="<%=userID %>">

<script language=javascript>
<%	Response.write "treeRoot = gFld(""ForumToc"", """ & FtypeName & """, ""CatList.asp?ItemID="& ItemID &"&CatID=0"")" & vbcrlf
	NowCount = 0

	SqlCom = "SELECT t.*, u.rights " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.dataParent=t.CtNodeID " _
		& " AND c.CtNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c JOIN CtUserSet AS cu ON cu.CtNodeID=c.CtNodeID " _
		& " AND cu.UserID=" & pkStr(userID,"") _
		& " WHERE c.dataParent=t.CtNodeID AND cu.rights IS NOT NULL) AS hasChild " _
		& " FROM CatTreeNode AS t LEFT JOIN CtUserSet AS u ON u.CtNodeID=t.CtNodeID " _
		& " AND u.UserID=" & pkStr(userID,"") _
		& " WHERE CtRootID = "& PkStr(ItemID,"") _
		&" Order by DataLevel, CatShowOrder"
	set RS = conn.execute(SqlCom)
	while not RS.eof
		xParent = "treeRoot"
		if rs(clevel) > 1 then	xParent = "N" & RS(cparent)
		if RS("CtNodeKind") = "C" then
			CatLink = "CatList.asp?ItemID="& ItemID &"&CatID="& rs("CtNodeID")
			CatLink = "#"
			if RS("hasChild")<> 0 then 
				Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs("hasChildFolder")&rs(cname)&rs("hasChild") &""", """& CatLink &"""))" & vbcrlf
			elseif RS("hasChildFolder") <> 0 then
				childCount = 0
				checkAllChild RS("CtNodeID")
				if childCount > 0 then
					Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs("hasChildFolder")&rs(cname)&rs("hasChild")&childCount &""", """& CatLink &"""))" & vbcrlf
				else
					Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs("hasChildFolder")&rs(cname)&rs("hasChild")&childCount &""", """& CatLink &"""))" & vbcrlf
				end if
			end if
		else
	     ForumLink = "ctXMLin.asp?ItemID="& ItemID &"&CtNodeID="& rs("CtNodeID")
	     ForumLink = "#"
	     
'	     sql1 = "select count(*) from CtUserSet where userID='" & userID & "' and CtNodeID=" & rs("CtNodeID")
'	     set rs1 = conn.execute(sql1)
'	     if cint(rs1(0)) = 0 then
	     if isNull(RS("rights")) then
	       ForumCheck = "<input type=checkbox name='CtNodeID' value='" & rs("CtNodeID") & "'>" & rs("CatName")
	     else
	       ForumCheck = "<input type=checkbox name='CtNodeID' value='" & rs("CtNodeID") & "' checked>" & rs("CatName")
	     end if
	     
	     if not isNull(RS("rights")) then _
	     	Response.Write "insDoc(N"& rs(cparent) &", gLnk(""ForumToc"", """& ForumCheck &""", """& ForumLink &"""))" & vbcrlf
		end if
		RS.moveNext
	wend


%>
     initializeDocument();
</script>
<HR>
<input name="submit" type="submit" value="確定"  onClick="return confirm('確定嗎？');">
</form>
</body></html>
<%
sub checkAllChild (xNodeID)
	response.write "// " & xNodeID & "==>" & childCount & vbCRLF
	xSql = "SELECT t.* " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.dataParent=t.CtNodeID " _
		& " AND c.CtNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c JOIN CtUserSet AS cu ON cu.CtNodeID=c.CtNodeID " _
		& " AND cu.UserID=" & pkStr(userID,"") _
		& " WHERE c.dataParent=t.CtNodeID AND cu.rights IS NOT NULL) AS hasChild " _
		& " FROM CatTreeNode AS t WHERE t.dataParent = "& xNodeID
'	response.write "// " & xNodeID & "==>" & xSql & vbCRLF
	set xRS = conn.execute(xSql)
	while not xRS.eof
		if xRS("hasChild") > 0 then	
			childCount = childCount + xRS("hasChild") 
		elseif xRS("hasChildFolder") > 0 then
			checkAllChild xRS("CtNodeID")
		end if
		
		xRS.moveNext
	wend
	response.write "// " & xNodeID & "==>" & childCount & vbCRLF
end sub
%>
