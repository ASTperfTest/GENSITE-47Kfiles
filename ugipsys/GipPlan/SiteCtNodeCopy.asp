<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<% 
dim CtRoot, orgRoot, dtLevel
	orgRoot = 33
	CtRoot = 2
	targetNode = 0
	dtLevel = 0 

'	response.write request.queryString
'	response.end
   
   	traverseTree 0, targetNode
   	
   	
   	
sub traverseTree (parent, np)
	SqlCom = "SELECT * FROM GipHy.dbo.CatTreeNode WHERE CtRootID = "& PkStr(orgRoot,"") _
		& " AND dataParent=" & parent & " Order by CatShowOrder"
	response.write sqlCom & "<HR>"
	set RS = conn.execute(SqlCom)
	myParent = cint(np)
	
	while not RS.eof
		response.write RS("CtNodeID") & RS("CatName") & "<BR>"
		sql = "INSERT INTO CatTreeNode(CtRootID, CtNodeKind, CatName, CtNameLogo, CatShowOrder, DataLevel" _
			& ", dataParent, CtUnitID, inUse, EditUserID, EditDate) VALUES(" _
			& dfn(CtRoot) & dfs(RS("CtNodeKind")) & dfs(RS("CatName")) & dfs(RS("CtNameLogo")) _
			& dfn(RS("CatShowOrder")) & dfn(RS("DataLevel")+dtLevel) & dfn(myParent) & dfn(RS("CtUnitID")) _
			& dfs(RS("inUse")) & dfs(session("userID")) & pkStr(date(),")")
		response.write sql & "<BR>"
		sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
		set RSx = conn.Execute(SQL)
		xNewIdentity = RSx(0)
'		response.write xNewIdentity & "<BR>"
		if RS("CtNodeKind") = "C" then 		traverseTree RS("CtNodeID"), xNewIdentity
		RS.moveNext
	wend
end sub

%>
<script language=vbs>
'	window.navigate "CatList.asp?ItemID=<%=CtRoot%>&CatID=<%=targetNode%>"
</script>

