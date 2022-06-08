<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GC1AP2" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link href="/css/tree.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<meta http-equiv="pragma: nocache">
<%
	if request.querystring("IType")="" then
		OpenCloseAnchor="Folder_toc.asp?IType=2"
		OpenCloseStr="展開"
		OpenCloseState="Close"
	else
		OpenCloseAnchor="Folder_toc.asp"
		OpenCloseStr="收合"
		OpenCloseState="Open"		
	end if
	FtypeName = "資料審稿"+"　　　<A href='"+OpenCloseAnchor+"'>"+OpenCloseStr+"</A>　<A href='GipApproveList.asp' target='ForumToc'>待審清單</A>"
	userId = session("userId")

   const cid     = 2	'CatID
   const cname   = 3	'CatName
   const cparent = 7	'DataParent
   const cChild  = 8	'ChildCount
   const clevel  = 6	'dataLevel


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
<script language=javascript>
var OpenCloseState = "<%=OpenCloseState%>"
<%	Response.write "treeRoot = gFld(""ForumToc"", """ & FtypeName & """, """")" & vbcrlf
	NowCount = 0
	
	SQLCom = "select CTR.* " & _
		"from CatTreeRoot CTR " & _
		"Where CTR.pvXdmp is not null " _
		& " AND (deptId IS NULL OR deptId LIKE N'" & session("deptId") & "%'" _
		& " OR '" & session("deptId") & "' LIKE deptId+'%')"
	set RStree = conn.execute(SqlCom)
	
	while not RStree.eof
		ItemID = RStree("ctRootId")
		session("ItemID")=ItemID
	
	
	SqlCom = "SELECT t.*, u.rights " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.dataParent=t.ctNodeId " _
		& " AND c.ctNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c JOIN CtUserSet2 AS cu ON cu.ctNodeId=c.ctNodeId " _
		& " AND cu.userId=" & pkStr(userId,"") _
		& " WHERE c.dataParent=t.ctNodeId AND cu.rights IS NOT NULL) AS hasChild " _
		& " FROM CatTreeNode AS t LEFT JOIN CtUserSet2 AS u ON u.ctNodeId=t.ctNodeId " _
		& " AND u.userId=" & pkStr(userId,"") _
		& " WHERE ctRootId = "& PkStr(ItemID,"") _
		&" Order by dataLevel, catShowOrder"
'	response.write "//" & sqlCom
'	response.end
	set RS = conn.execute(SqlCom)
	if not RS.eof then
		xParent = "treeRoot"
		CatLink = "blank.htm" 'CatList.asp?ItemID="& ItemID &"&CatID="& rs("ctNodeId")
'		if RS("hasChild")<> 0 OR RS("hasChildFolder") <> 0 then 
			Response.Write "T"& itemID &"= insFld(" & xParent &", gFld(""ForumToc"", """& RStree("CtRootName") &""", """& CatLink &"""))" & vbcrlf
'		end if
	end if	
	while not RS.eof
		xParent = "T" &itemID
		if rs(clevel) > 1 then	xParent = "N" & RS(cparent)
		if RS("ctNodeKind") = "C" then
			CatLink = "blank.htm" 
			if RS("hasChild")<> 0 then 
				Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
			elseif RS("hasChildFolder") <> 0 then
				childCount = 0
				checkAllChild RS("ctNodeId")
				if childCount > 0 then _
					Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
			end if
		else
	     ForumLink = "GipApproveList.asp?ctRootId="& ItemID &"&ctNodeId="& rs("ctNodeId")
	     if not isNull(RS("rights")) then
	     	Response.Write "insDoc("& xParent &", gLnk(""ForumToc"", """& rs("CatName") &""", """& ForumLink &"""))" & vbcrlf
		  end if
		end if
		RS.moveNext
	wend

	RStree.moveNext
wend		
	
%>
if (OpenCloseState == "Close")
     initializeDocument();
else
     initializeDocumentClose();
</script>
</body></html>
<%
sub checkAllChild (xNodeID)
	xSql = "SELECT t.* " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.dataParent=t.ctNodeId " _
		& " AND c.ctNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c JOIN CtUserSet2 AS cu ON cu.ctNodeId=c.ctNodeId " _
		& " AND cu.userId=" & pkStr(userId,"") _
		& " WHERE c.dataParent=t.ctNodeId AND cu.rights IS NOT NULL) AS hasChild " _
		& " FROM CatTreeNode AS t WHERE t.dataParent = "& xNodeID
	set xRS = conn.execute(xSql)
	while not xRS.eof
		if xRS("hasChild") > 0 then	
			childCount = childCount + xRS("hasChild") 
		elseif xRS("hasChildFolder") > 0 then
			checkAllChild xRS("ctNodeId")
		end if
		
		xRS.moveNext
	wend
end sub
%>