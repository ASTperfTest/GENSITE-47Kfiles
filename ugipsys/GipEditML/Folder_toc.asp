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
	if request.querystring("IType")="" then
		OpenCloseAnchor="Folder_toc.asp?IType=2"
		OpenCloseStr="展開"
		OpenCloseState="Close"
	else
		OpenCloseAnchor="Folder_toc.asp"
		OpenCloseStr="收合"
		OpenCloseState="Open"		
	end if
	FtypeName = "資料上稿"+"　　　<A href='"+OpenCloseAnchor+"'>"+OpenCloseStr+"</A>"
	userID = session("userID")

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
<script language=javascript>
var OpenCloseState = "<%=OpenCloseState%>"
<%	Response.write "treeRoot = gFld(""ForumToc"", """ & FtypeName & """, """")" & vbcrlf
	NowCount = 0

	SQLCom = "select * from CatTreeRoot Where pvXdmp is not null " _
		& " AND (deptID IS NULL OR deptID LIKE '" & session("deptID") & "%'" _
		& " OR '" & session("deptID") & "' LIKE deptID+'%')"
	set RStree = conn.execute(SqlCom)
	
	while not RStree.eof
		ItemID = RStree("CtRootID")
	session("ItemID")=ItemID


  if Instr(session("uGrpID")&",", "HTSD,") > 0 then  
	SqlCom = "SELECT t.*, 'Y' AS Rights" _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.DataParent=t.CtNodeID " _
		& " AND c.CtNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.DataParent=t.CtNodeID) AS hasChild " _
		& " FROM CatTreeNode AS t " _
		& " WHERE CtRootID = "& PkStr(ItemID,"") _
		&" Order by DataLevel, CatShowOrder"
  else
	SqlCom = "SELECT t.*, u.Rights " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.DataParent=t.CtNodeID " _
		& " AND c.CtNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c JOIN CtUserSet AS cu ON cu.CtNodeID=c.CtNodeID " _
		& " AND cu.UserID=" & pkStr(userID,"") _
		& " WHERE c.DataParent=t.CtNodeID AND cu.Rights IS NOT NULL) AS hasChild " _
		& " FROM CatTreeNode AS t LEFT JOIN CtUserSet AS u ON u.CtNodeID=t.CtNodeID " _
		& " AND u.UserID=" & pkStr(userID,"") _
		& " WHERE CtRootID = "& PkStr(ItemID,"") _
		&" Order by DataLevel, CatShowOrder"
	xSql = "SELECT count(*) FROM CtUserSet AS u JOIN CatTreeNode AS t ON u.CtNodeID=t.CtNodeID" _
		& " AND u.UserID=" & pkStr(userID,"") _
		& " WHERE CtRootID = "& PkStr(ItemID,"")
	set RS= conn.execute(xSql)
	if RS(0) = 0 then	SqlCom = "SELECT * FROM AP WHERE 1=2"
	
	
  end if 
'	response.write "//" & sqlCom
'	response.end
	set RS = conn.execute(SqlCom)
	if not RS.eof then
		xParent = "treeRoot"
		CatLink = "blank.htm" 'CatList.asp?ItemID="& ItemID &"&CatID="& rs("CtNodeID")
'		if RS("hasChild")<> 0 OR RS("hasChildFolder") <> 0 then 
			Response.Write "T"& itemID &"= insFld(" & xParent &", gFld(""ForumToc"", """& RStree("CtRootName") &""", """& CatLink &"""))" & vbcrlf
'		end if
	end if
	while not RS.eof
		xParent = "T" &itemID
		if rs(clevel) > 1 then	xParent = "N" & RS(cparent)
		if RS("CtNodeKind") = "C" then
			CatLink = "blank.htm" 'CatList.asp?ItemID="& ItemID &"&CatID="& rs("CtNodeID")
			if RS("hasChild")<> 0 then 
				Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
			elseif RS("hasChildFolder") <> 0 then
				childCount = 0
				checkAllChild RS("CtNodeID")
				if childCount > 0 then _
					Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
			end if
		else
	     ForumLink = "ctXMLin.asp?ItemID="& ItemID &"&CtNodeID="& rs("CtNodeID")
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
  if Instr(session("uGrpID")&",", "HTSD,") > 0 then  
	xSql = "SELECT t.*, 'Y' AS Rights " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.DataParent=t.CtNodeID " _
		& " AND c.CtNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.DataParent=t.CtNodeID) AS hasChild " _
		& " FROM CatTreeNode AS t WHERE t.DataParent = "& xNodeID
  else
	xSql = "SELECT t.* " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.DataParent=t.CtNodeID " _
		& " AND c.CtNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c JOIN CtUserSet AS cu ON cu.CtNodeID=c.CtNodeID " _
		& " AND cu.UserID=" & pkStr(userID,"") _
		& " WHERE c.DataParent=t.CtNodeID AND cu.Rights IS NOT NULL) AS hasChild " _
		& " FROM CatTreeNode AS t WHERE t.DataParent = "& xNodeID
  end if 
	set xRS = conn.execute(xSql)
	while not xRS.eof
		if xRS("hasChild") > 0 then	
			childCount = childCount + xRS("hasChild") 
		elseif xRS("hasChildFolder") > 0 then
			checkAllChild xRS("CtNodeID")
		end if
		
		xRS.moveNext
	wend
end sub
%>
