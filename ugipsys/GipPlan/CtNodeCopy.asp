<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<% 
dim CtRoot, orgRoot, dtLevel
	orgRoot = request("orgTree")
	CtRoot = request("ItemID")
	targetNode = request("CatID")
	dtLevel = cint(request("dataLevel")) - 1

'	response.write request.queryString
'	response.end
   
   	traverseTree 0, targetNode
   	
   	
   	
sub traverseTree (parent, np)
	SqlCom = "SELECT * FROM CatTreeNode WHERE ctRootId = "& PkStr(orgRoot,"") _
		& " AND dataParent=" & parent & " Order by catShowOrder"
	response.write sqlCom & "<HR>"
	set RS = conn.execute(SqlCom)
	myParent = cint(np)
	
	while not RS.eof
		response.write RS("CtNodeID") & RS("catName") & "<BR>"
		sql = "INSERT INTO CatTreeNode(ctRootId, ctNodeKind, catName, ctNameLogo, catShowOrder, dataLevel" _
			& ", dataParent, ctUnitId, inUse, editUserId, editDate) VALUES(" _
			& dfn(CtRoot) & dfs(RS("ctNodeKind")) & dfs(RS("catName")) & dfs(RS("ctNameLogo")) _
			& dfn(RS("catShowOrder")) & dfn(RS("dataLevel")+dtLevel) & dfn(myParent) & dfn(RS("ctUnitId")) _
			& dfs(RS("inUse")) & dfs(session("userID")) & pkStr(date(),")")
		response.write sql & "<BR>"
		sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
		set RSx = conn.Execute(SQL)
		xNewIdentity = RSx(0)
'		response.write xNewIdentity & "<BR>"
		if RS("ctNodeKind") = "C" then 		traverseTree RS("CtNodeID"), xNewIdentity
		RS.moveNext
	wend
end sub

%>
<script language=vbs>
	window.navigate "CatList.asp?ItemID=<%=CtRoot%>&CatID=<%=targetNode%>"
</script>

