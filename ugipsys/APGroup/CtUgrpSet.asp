<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "HT001" 
' ============= Modified by Chris, 2006/09/12, to handle 上稿權限能以使用者設定 聯集 群組設定 ========================'
'		Document: 950912_智庫GIP擴充.doc
'  modified list:
'	新增此程式
%>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"; cache=no-caches;>
<meta http-equiv="pragma: nocache">
<%


	FtypeName = "資料上稿"


   const cid     = 2	'CatID
   const cname   = 3	'CatName
   const cparent = 7	'DataParent
   const cChild  = 8	'ChildCount
   const clevel  = 6	'dataLevel


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
  ugrpID = request("ugrpID")
  aType = request("aType")
%>
<form name="form1" method="post" action="CtUgrpSet_act.asp">
  <input type="hidden" name="ugrpID" value="<%=ugrpID %>">
  <input type="hidden" name="aType" value="<%=aType%>">

<script language=javascript>
<%	Response.write "treeRoot = gFld(""ForumToc"", """ & FtypeName & """, ""CatList.asp?ItemID="& ItemID &"&CatID=0"")" & vbcrlf
	NowCount = 0

  SQLCom = "select * from CatTreeRoot Where pvXdmp is not null " _
  	& " AND (deptId IS NULL OR deptId LIKE '" & session("deptId") & "%')" 
  set RStree = conn.execute(SqlCom)
  
  while not RStree.eof
  	ItemID = RStree("ctRootId")
	Response.Write "T"& itemID &"= insFld(treeRoot, gFld(""ForumToc"", """& RStree("CtRootName") &""", ""blank.htm""))" & vbcrlf

	SqlCom = "SELECT n.*, g.rights FROM CatTreeNode AS n LEFT JOIN CtUgrpSet AS g ON g.ctNodeId=n.ctNodeId" _
		& " AND g.ugrpId = " & pkStr(ugrpID,"") _
		& " AND g.aType=" & pkStr(aType,"") _
		& " WHERE ctRootId = "& PkStr(ItemID,"") &" Order by dataLevel, catShowOrder"
	set RS = conn.execute(SqlCom)
	while not RS.eof
		xParent = "T" &itemID
		if rs(clevel) > 1 then	xParent = "N" & RS(cparent)
		if RS("CtNodeKind") = "C" then
			CatLink = "CatList.asp?ItemID="& ItemID &"&CatID="& rs("ctNodeId")
			CatLink = "#"
			Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
		else
	     ForumLink = "ctXMLin.asp?ItemID="& ItemID &"&ctNodeId="& rs("ctNodeId")
	     ForumLink = "#"
	     
	     if rs("rights")> 0 then
	       ForumCheck = "<input type=checkbox name='ctNodeId' value='" & rs("ctNodeId") & "' checked>" & rs("CatName")
	     else
	       ForumCheck = "<input type=checkbox name='ctNodeId' value='" & rs("ctNodeId") & "'>" & rs("CatName")
	     end if
'	     sql1 = "select count(*) from CtUserSet where userId=N'" & userId & "' and ctNodeId=" & rs("ctNodeId")
'	     set rs1 = conn.execute(sql1)
'	     if cint(rs1(0)) = 0 then
'	       ForumCheck = "<input type=checkbox name='ctNodeId' value='" & rs("ctNodeId") & "'>" & rs("CatName")
'	     else
'	       ForumCheck = "<input type=checkbox name='ctNodeId' value='" & rs("ctNodeId") & "' checked>" & rs("CatName")
'	     end if
	     
	     Response.Write "insDoc("& xParent &", gLnk(""ForumToc"", """& ForumCheck &""", """& ForumLink &"""))" & vbcrlf
		end if
		RS.moveNext
	wend
	RStree.moveNext
  wend

%>
     initializeDocument();
</script>
<HR>
<input name="submit" type="submit" value="確定"  onClick="return confirm('確定嗎？');">
</form>
</body></html>
