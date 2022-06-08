<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<html><head>
<meta http-equiv="pragma: nocache">
<%


	FtypeName = "資料上稿"


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
<form name="form1" method="post">
  <input type="hidden" name="userID" value="<%=userID %>">

<script language=javascript>
<%	Response.write "treeRoot = gFld(""ForumToc"", """ & FtypeName & """, ""CatList.asp?ItemID="& ItemID &"&CatID=0"")" & vbcrlf
	NowCount = 0

  SQLCom = "select * from CatTreeRoot Where pvXdmp is not null " _
  	& " AND (deptID IS NULL OR deptID LIKE N'" & session("deptID") & "%')" 
  set RStree = conn.execute(SqlCom)
  
  while not RStree.eof
  	ItemID = RStree("CtRootID")
	Response.Write "T"& itemID &"= insFld(treeRoot, gFld(""ForumToc"", """& RStree("CtRootName") &""", ""blank.htm""))" & vbcrlf

	SqlCom = "SELECT * FROM CatTreeNode WHERE CtRootID = "& PkStr(ItemID,"") &" Order by DataLevel, CatShowOrder"
	set RS = conn.execute(SqlCom)
	while not RS.eof
		xParent = "T" &itemID
		if rs(clevel) > 1 then	xParent = "N" & RS(cparent)
		if RS("CtNodeKind") = "C" then
			CatLink = "CatList.asp?ItemID="& ItemID &"&CatID="& rs("CtNodeID")
			CatLink = "#"
			Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
		else
	     ForumLink = "ctXMLin.asp?ItemID="& ItemID &"&CtNodeID="& rs("CtNodeID")
	     ForumLink = "#"
	     
	     sql1 = "select count(*) from CtUserSet where userID=N'" & userID & "' and CtNodeID=" & rs("CtNodeID")
	     set rs1 = conn.execute(sql1)
	     if cint(rs1(0)) = 0 then
	       ForumCheck = "<input type=checkbox name='CtNodeID' value='" & rs("CtNodeID") & "'>" & rs("CatName")
	     else
	       ForumCheck = "<input type=checkbox name='CtNodeID' value='" & rs("CtNodeID") & "' checked>" & rs("CatName")
	     end if
	     
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
<input name="submit" type="button" value="確定"  onClick="VBS: doPost()">
</form>
<script language=VBS>
sub doPost()
'	alert document.all("ChStuNo"&NoChoice).value&";;;"&document.all("ChStuName"&NoChoice).value
	window.opener.xReturn "AAA,BBB,"&";;;"&"aaa,bbb,"
'	window.opener.xReturn document.all("ChStuNo"&NoChoice).value&";;;"&document.all("ChStuName"&NoChoice).value
'	window.returnValue=xID & ";;;" & xName
	window.close()
	
end sub
</script>                               
</body></html>
