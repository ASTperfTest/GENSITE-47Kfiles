<%@ CodePage = 65001 %><% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#Include file = "inc/index.inc" -->
<!--#Include file = "inc/client.inc" -->
<!--#Include file = "inc/dbFunc.inc" -->
<html>
<head>
<meta http-equiv="pragma: nocache">
<meta http-equiv="Content-Language" content="zh-tw">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>財政部全球資訊網動態資料庫網站維護帳號申請</title>
</head>

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
<p align="center"><font color="#0000FF">財政部全球資訊網動態資料庫網站維護帳號申請</font></p>
<p align="center">請填選後印出，主管核章後傳真資訊小組(23517147)，本小組據以開設帳號及上稿權限</p>
<p>單&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 位：<select size="1" name="D1">
			<%SQL="Select deptID,deptName from dept WHERE inUse='Y' Order by kind, deptID"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select></p>
<p>姓　　名：<input type="text" name="T1" size="17">&nbsp;&nbsp;&nbsp;&nbsp;EMAIL：<input type="text" name="T1" size="17">&nbsp;&nbsp;&nbsp;&nbsp; 
聯絡電話：<input type="text" name="T1" size="17"></p>
<p>登入帳號：<input type="text" name="T1" size="17">&nbsp;&nbsp;&nbsp;&nbsp;密　碼：<input type="text" name="T1" size="17">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;密碼確認：<input type="text" name="T1" size="17"></p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
申請單位主管：　　　　　　申請人：</p>
<p>申請項目（點選展開, 可複選） </p>
<HR>
<TABLE border=0 align="center">
<TR><TD>
<script language=javascript>
<%	Response.write "treeRoot = gFld(""ForumToc"", """ & FtypeName & """, ""CatList.asp?ItemID="& ItemID &"&CatID=0"")" & vbcrlf
	NowCount = 0

	SqlCom = "SELECT n.*, u.ctUnitKind FROM CatTreeNode AS n LEFT JOIN CtUnit AS u ON u.CtUnitID=n.CtUnitID WHERE CtRootID = "& PkStr(ItemID,"") &" Order by DataLevel, CatShowOrder"
	set RS = conn.execute(SqlCom)
	while not RS.eof
		xParent = "treeRoot"
		if rs(clevel) > 1 then	xParent = "N" & RS(cparent)
		if RS("CtNodeKind") = "C" then
			CatLink = "CatList.asp?ItemID="& ItemID &"&CatID="& rs("CtNodeID")
			CatLink = ""
			Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
		elseif rs("ctUnitKind") <> "U" then
	     ForumLink = "ctXMLin.asp?ItemID="& ItemID &"&CtNodeID="& rs("CtNodeID")
	     ForumLink = ""
	     
	       ForumCheck = "<input type=checkbox name='CtNodeID' value='" & rs("CtNodeID") & "'>" & rs("CatName")
	     
	     Response.Write "insDoc(N"& rs(cparent) &", gLnk(""ForumToc"", """& ForumCheck &""", """& ForumLink &"""))" & vbcrlf
		end if
		RS.moveNext
	wend


%>
     initializeDocument();
</script>
</TD></TR></TABLE>
<HR>
</body></html>