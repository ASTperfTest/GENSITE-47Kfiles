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
<%
	xShowType=request.querystring("showType")
	if xShowType="4" then
		xShowTypeStr="複製"
	else
		xShowTypeStr="參照"	
	end if
%>
<html><head>
<meta http-equiv="pragma: nocache">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<div id="FuncName">
	<h1><%=Title%>/多向出版(引用方式: <%=xShowTypeStr%>)</h1>
	<div id="Nav">
	       <a href="Javascript:window.navigate('DsdXMLEdit.asp?iCuItem=<%=request.querystring("iCuItem")%>');">回編修</a>	    
	</div>
	<div id="ClearFloat"></div>
</div>
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

  SQLCom = "select * from CatTreeRoot Where pvXdmp is not null " _
  	& " AND (deptID IS NULL OR deptID LIKE '" & session("deptID") & "%')" 
  set RStree = conn.execute(SqlCom)
'	ItemID = session("exRootID")		'-- 要處理哪個 tree
  	ItemID = RStree("CtRootID")
	SQLCom = "select * from CatTreeRoot Where CtRootID = "& ItemID 
	set RS = conn.execute(SqlCom)
	
	if request.querystring("IType")="" then
		OpenCloseAnchor="DsdXMLAdd_Multi.asp?IType=2&showType="&request.querystring("showType")&"&iCuItem="&request.querystring("iCuItem")
		OpenCloseStr="展開"
		OpenCloseState="Close"
	else
		OpenCloseAnchor="DsdXMLAdd_Multi.asp?showType="&request.querystring("showType")&"&iCuItem="&request.querystring("iCuItem")
		OpenCloseStr="收合"
		OpenCloseState="Open"		
	end if
	if xShowType="4" then
		FtypeName = RS("CtRootName")+"<A href='#'></A>　　　　　<A href='"+OpenCloseAnchor+"'>"+OpenCloseStr+"</A>　<A href='Javascript: formSubmit();'>確定新增</A>"
	else
		FtypeName = RS("CtRootName")+"<A href='#'></A>　　　　　<A href='"+OpenCloseAnchor+"'>"+OpenCloseStr+"</A>　<A href='Javascript: formSubmit();'>確定新增</A>　<A href='Javascript: refList();'>參照清單</A>"
	end if
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
<form name="form1" method="post" action="DsdXMLAdd_Multi_act.asp?iCuItem=<%=request.querystring("iCuItem")%>">
<input name=showType type="hidden" value="<%=xShowType%>">

<%
    userID = session("userID")
    session("DsdXMLAdd_Multi_List")=""
    '----給參照清單用的SQL字串
    if xShowType="5" then
    	SQLRef5List = "Select t.CatName,CDT.sTitle,CDT.dEditDate,D.deptName,CDT.iCuItem " _
		& " FROM CatTreeNode AS t " _
		& " LEFT JOIN CtUserSet AS u ON u.CtNodeID=t.CtNodeID " _
		& " Left Join CuDTGeneric CDT ON t.CtUnitID=CDT.iCTUnit AND CDT.ShowType=N'"&xShowType&"' AND CDT.refID="&request.querystring("iCuItem") _
		& " Left Join CtUnit CU ON CDT.iCTUnit=CU.CtUnitID " _
		& " Left Join dept D On CDT.iDept=D.deptID " _
		& " WHERE CtRootID = "& PkStr(ItemID,"") & " AND u.UserID=" & pkStr(userID,"") & " AND CDT.iCuItem is not null" _
		& " Order by CDT.dEditDate DESC"
    	session("DsdXMLAdd_Multi_List")=SQLRef5List
    end if
%>
  <input type="hidden" name="userID" value="<%=userID %>">

<script language=javascript>
var OpenCloseState = "<%=OpenCloseState%>"
<%	Response.write "treeRoot = gFld(""ForumToc"", """ & FtypeName & """, ""#"")" & vbcrlf
	NowCount = 0

	SqlCom = "SELECT t.*, u.Rights,CU.ctUnitKind,CU.CtUnitID,b.sBaseDSDXML " _
		& ", (SELECT count(*) FROM CatTreeNode AS c WHERE c.DataParent=t.CtNodeID " _
		& " AND c.CtNodeKind='C') AS hasChildFolder " _
		& ", (SELECT count(*) FROM CatTreeNode AS c JOIN CtUserSet AS cu ON cu.CtNodeID=c.CtNodeID " _
		& " AND cu.UserID=" & pkStr(userID,"") _
		& " WHERE c.DataParent=t.CtNodeID AND cu.Rights IS NOT NULL) AS hasChild " _
		& ", (SELECT count(iCUItem) FROM CatTreeNode AS c Left Join CuDTGeneric CDT ON c.CtUnitID=CDT.iCTUnit " _
		& " WHERE CDT.ShowType=N'"&xShowType&"' AND c.CtNodeID=t.CtNodeID AND CDT.refID="&request.querystring("iCuItem")&") AS refID " _
		& " FROM CatTreeNode AS t LEFT JOIN CtUserSet AS u ON u.CtNodeID=t.CtNodeID " _
		& " AND u.UserID=" & pkStr(userID,"") _
		& " Left Join CtUnit CU ON t.CtUnitID=CU.CtUnitID " _
		& " Left Join BaseDSD b ON b.iBaseDSD=CU.iBaseDSD " _
		& " WHERE CtRootID = "& PkStr(ItemID,"") _
		&" Order by DataLevel, CatShowOrder"
'response.write sqlcom
'response.end		
	set RS = conn.execute(SqlCom)
	while not RS.eof
		xParent = "treeRoot"
		if rs(clevel) > 1 then	xParent = "N" & RS(cparent)
		if RS("CtNodeKind") = "C" then
			CatLink = "#"
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
	     ForumLink = "#"
	     if not isNull(RS("rights")) then
	        if RS("ctUnitKind")="U" or isNull(RS("CtUnitID")) or not isNull(RS("sBaseDSDXML")) then
	        	ForumCheck = rs("CatName")	     
	     	elseif RS("refID")>=1 then
	        	ForumCheck = "<input type=checkbox name='CtNodeID' value='" & cStr(rs("CtNodeID")) & "' checked disabled>" & rs("CatName")	     
	     	else
	        	ForumCheck = "<input type=checkbox name='CtNodeID' value='" & cStr(rs("CtNodeID")) & "'>" & rs("CatName")	     
	        end if
	     	Response.Write "insDoc("& xParent &", gLnk(""ForumToc"", """& ForumCheck &""", """& ForumLink &"""))" & vbcrlf
		  end if
		end if
		RS.moveNext
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
sub refList()
	window.navigate "DsdXMLAdd_Multi_List.asp?<%=request.querystring%>"
end sub
</script>
