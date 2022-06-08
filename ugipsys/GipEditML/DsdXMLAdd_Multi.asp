<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="新增"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXML"
%>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
	xshowType=request.querystring("showType")
	if xshowType="4" then
		xshowTypeStr="複製"
	else
		xshowTypeStr="參照"	
	end if
%>
<html><head>
<meta http-equiv="pragma: nocache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<div id="FuncName">
	<h1><%=Title%>/多向出版(引用方式: <%=xshowTypeStr%>)</h1>
	<div id="Nav">
	       <a href="Javascript:window.navigate('DsdXMLEdit.asp?icuitem=<%=request.querystring("icuitem")%>');">回編修</a>	    
	</div>
	<div id="ClearFloat"></div>
</div>
<%
sub checkAllChild (xNodeID)
	xSql = "SELECT t.* " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.dataParent=t.ctNodeId " _
		& " AND c.ctNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c JOIN CtUserSet AS cu ON cu.ctNodeId=c.ctNodeId " _
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

  SQLCom = "select * from CatTreeRoot Where pvXdmp is not null " _
  	& " AND (deptId IS NULL OR deptId LIKE '" & session("deptId") & "%')" 
  set RStree = conn.execute(SqlCom)
'	ItemID = session("exRootID")		'-- 要處理哪個 tree
  	ItemID = RStree("ctRootId")
	SQLCom = "select * from CatTreeRoot Where ctRootId = "& ItemID 
	set RS = conn.execute(SqlCom)
	
	if request.querystring("IType")="" then
		OpenCloseAnchor="DsdXMLAdd_Multi.asp?IType=2&showType="&request.querystring("showType")&"&icuitem="&request.querystring("icuitem")
		OpenCloseStr="展開"
		OpenCloseState="Close"
	else
		OpenCloseAnchor="DsdXMLAdd_Multi.asp?showType="&request.querystring("showType")&"&icuitem="&request.querystring("icuitem")
		OpenCloseStr="收合"
		OpenCloseState="Open"		
	end if
	convertsimAnchor = ""
	if checkGIPconfig("convertsim") then
		convertsimAnchor = "　<A href='Javascript: formSubmitSim();'>確定新增(轉簡體)</A>"
	end if
	if xshowType="4" then
		FtypeName = "資料上稿" +"<A href='#'></A>　　　　　<A href='"+OpenCloseAnchor+"'>"+OpenCloseStr+"</A>　<A href='Javascript: formSubmit();'>確定新增</A>" + convertsimAnchor
	else
		FtypeName = "資料上稿" +"<A href='#'></A>　　　　　<A href='"+OpenCloseAnchor+"'>"+OpenCloseStr+"</A>　<A href='Javascript: formSubmit();'>確定新增</A>" + convertsimAnchor + "　<A href='Javascript: refList();'>參照清單</A>"
	end if
	userId = session("userId")
	session("ItemID")=ItemID

   const cid     = 2	'CatID
   const cname   = 3	'CatName
   const cparent = 7	'dataParent
   const cChild  = 8	'ChildCount
   const clevel  = 6	'dataLevel


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
<form name="form1" method="post" action="DsdXMLAdd_Multi_act.asp?icuitem=<%=request.querystring("icuitem")%>">
<input name=showType type="hidden" value="<%=xshowType%>">

<%
    userId = session("userId")
    session("DsdXMLAdd_Multi_List")=""
    '----給參照清單用的SQL字串
    if xshowType="5" then
    	SQLRef5List = "Select t.catName,CDT.stitle,CDT.deditDate,D.deptName,CDT.icuitem " _
		& " FROM CatTreeNode AS t " _
		& " LEFT JOIN CtUserSet AS u ON u.ctNodeId=t.ctNodeId " _
		& " Left Join CuDtGeneric CDT ON t.ctUnitId=CDT.ictunit AND CDT.showType=N'"&xshowType&"' AND CDT.refId="&request.querystring("icuitem") _
		& " Left Join CtUnit CU ON CDT.ictunit=CU.ctUnitId " _
		& " Left Join Dept D On CDT.idept=D.deptId " _
		& " WHERE ctRootId = "& PkStr(ItemID,"") & " AND u.userId=" & pkStr(userId,"") & " AND CDT.icuitem is not null" _
		& " Order by CDT.deditDate DESC"
    	session("DsdXMLAdd_Multi_List")=SQLRef5List
    end if
%>
  <input type="hidden" name="userId" value="<%=userId %>">

<script language=javascript>
var OpenCloseState = "<%=OpenCloseState%>"
<%	Response.write "treeRoot = gFld(""ForumToc"", """ & FtypeName & """, ""#"")" & vbcrlf
	NowCount = 0

SQLCom = "select * from CatTreeRoot Where pvXdmp is not null " _
	& " AND (deptId IS NULL OR deptId LIKE '" & session("deptId") & "%')" 
set RStree = conn.execute(SqlCom)
	
while not RStree.eof
		ItemID = RStree("ctRootId")
		session("ItemID")=ItemID

	SqlCom = "SELECT t.*, u.rights,CU.ctUnitKind,CU.ctUnitId,b.sbaseDsdxml " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.dataParent=t.ctNodeId " _
		& " AND c.ctNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c JOIN CtUserSet AS cu ON cu.ctNodeId=c.ctNodeId " _
		& " AND cu.userId=" & pkStr(userId,"") _
		& " WHERE c.dataParent=t.ctNodeId AND cu.rights IS NOT NULL) AS hasChild " _
		& ", (SELECT count(icuitem) FROM CatTreeNode AS c Left Join CuDtGeneric CDT ON c.ctUnitId=CDT.ictunit " _
		& " WHERE CDT.showType=N'"&xshowType&"' AND c.ctNodeId=t.ctNodeId AND CDT.refId="&request.querystring("icuitem")&") AS refId " _
		& " FROM CatTreeNode AS t LEFT JOIN CtUserSet AS u ON u.ctNodeId=t.ctNodeId " _
		& " AND u.userId=" & pkStr(userId,"") _
		& " Left Join CtUnit CU ON t.ctUnitId=CU.ctUnitId " _
		& " Left Join BaseDsd b ON b.ibaseDsd=CU.ibaseDsd " _
		& " WHERE ctRootId = "& PkStr(ItemID,"") _
		&" Order by dataLevel, catShowOrder"
'response.write sqlcom
'response.end		
	set RS = conn.execute(SqlCom)
	if not RS.eof then
		xParent = "treeRoot"
		CatLink = "" 'CatList.asp?ItemID="& ItemID &"&CatID="& rs("ctNodeId")
'		if RS("hasChild")<> 0 OR RS("hasChildFolder") <> 0 then 
			Response.Write "T"& itemID &"= insFld(" & xParent &", gFld(""ForumToc"", """& RStree("CtRootName") &""", """& CatLink &"""))" & vbcrlf
'		end if
	end if
	while not RS.eof
		xParent = "T" &itemID
		if rs(clevel) > 1 then	xParent = "N" & RS(cparent)
		if RS("ctNodeKind") = "C" then
			CatLink = "#"
			if RS("hasChild")<> 0 then 
				Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
			elseif RS("hasChildFolder") <> 0 then
				childCount = 0
				checkAllChild RS("ctNodeId")
				if childCount > 0 then _
					Response.Write "N"& rs(cid) &"= insFld(" & xParent &", gFld(""ForumToc"", """& rs(cname) &""", """& CatLink &"""))" & vbcrlf
			end if
		else
	     ForumLink = "ctXMLin.asp?ItemID="& ItemID &"&ctNodeId="& rs("ctNodeId")
	     ForumLink = "#"
	     if not isNull(RS("rights")) then
	        if RS("ctUnitKind")="U" or isNull(RS("ctUnitId")) or not isNull(RS("sbaseDsdxml")) then
	        	ForumCheck = rs("catName")	     
	     	elseif RS("refId")>=1 then
	        	ForumCheck = "<input type=checkbox name='ctNodeId' value='" & cStr(rs("ctNodeId")) & "' checked disabled>" & rs("catName")	     
	     	else
	        	ForumCheck = "<input type=checkbox name='ctNodeId' value='" & cStr(rs("ctNodeId")) & "'>" & rs("catName")	     
	        end if
	     	Response.Write "insDoc("& xParent &", gLnk(""ForumToc"", """& ForumCheck &""", """& ForumLink &"""))" & vbcrlf
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
</form>
</body></html>
<script language=VBS>
sub formSubmit()
	form1.submit
end sub
sub formSubmitSim()
	chky=msgbox("注意！"& vbcrlf & vbcrlf & "確定轉為簡體嗎?" & vbcrlf , 48+1, "請注意！！")
        if chky=vbok then	
		form1.action = "DsdXMLAdd_Multi_act.asp?sim=Y&icuitem=<%=request.querystring("icuitem")%>"
		form1.submit
	end if
end sub
sub refList()
	window.navigate "DsdXMLAdd_Multi_List.asp?<%=request.querystring%>"
end sub
</script>
