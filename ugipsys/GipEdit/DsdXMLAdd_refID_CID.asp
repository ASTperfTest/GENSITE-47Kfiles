<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
HTProgCode="GC1AP1"
%>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<html><head>
<meta http-equiv="pragma: nocache">
<%
sub checkAllChild (xNodeID)
	xSql = "SELECT t.* " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.DataParent=t.CtNodeID " _
		& " AND c.CtNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c JOIN CtUserSet AS cu ON cu.CtNodeID=c.CtNodeID " _
		& " AND cu.UserID=" & pkStr(userID,"") _
		& " WHERE c.DataParent=t.CtNodeID AND cu.Rights IS NOT NULL) AS hasChild " _
		& " FROM CatTreeNode AS t WHERE t.DataParent = "& xNodeID
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

	ItemID = session("exRootID")		'-- 要處理哪個 tree
	SQLCom = "select * from CatTreeRoot Where CtRootID = "& ItemID 
	set RS = conn.execute(SqlCom)


	FtypeName = RS("CtRootName")
	userID = session("userID")
	session("ItemID")=ItemID

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
<script src="FolderTree.js"></script>
<title></title>
<style>
.FolderStyle { text-decoration: none; color:#000000}
</style>
</head>
<body leftmargin="3" topmargin="3">
<form name="form1" method="post" action="DsdXMLAdd_refID_CID_act.asp">

選取引用資料條件 - 點選主題單元
<input name="submit" type="submit" value="確定">
<HR>

<%
  userID = session("userID")
%>
  <input type="hidden" name="userID" value="<%=userID %>">

<script language=javascript>
<%	Response.write "treeRoot = gFld(""_self"", """ & FtypeName & """, ""#"")" & vbcrlf
	NowCount = 0

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
	set RS = conn.execute(SqlCom)
	while not RS.eof
		xParent = "treeRoot"
		if rs(clevel) > 1 then	xParent = "N" & RS(cparent)
		if RS("CtNodeKind") = "C" then
			CatLink = "#"
			if RS("hasChild")<> 0 then 
				Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""_self"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
			elseif RS("hasChildFolder") <> 0 then
				childCount = 0
				checkAllChild RS("CtNodeID")
				if childCount > 0 then _
					Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""_self"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
			end if
		else
	     ForumLink = "ctXMLin.asp?ItemID="& ItemID &"&CtNodeID="& rs("CtNodeID")
	     ForumLink = "#"
	     if not isNull(RS("rights")) then
	        ForumCheck = "<input type=checkbox name='CtNodeID' value='" & cStr(rs("CtNodeID"))+"|||"+rs("CatName") & "'>" & rs("CatName")	     
	     	Response.Write "insDoc("& xParent &", gLnk(""_self"", """& ForumCheck &""", """& ForumLink &"""))" & vbcrlf
		  end if
		end if
		RS.moveNext
	wend

%>
     initializeDocumentClose();
</script>
</form>
</body></html>
