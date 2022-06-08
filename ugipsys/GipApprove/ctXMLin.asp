<%@ CodePage = 65001 %>
<% 
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="資料審稿"
HTProgCode="GC1AP2"
HTProgPrefix="DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%	

	sql = "SELECT u.*,b.sBaseTableName FROM CatTreeNode AS n LEFT JOIN CtUnit AS u ON u.CtUnitID=n.CtUnitID" _
		& " Left Join BaseDSD As b ON u.iBaseDSD=b.iBaseDSD" _
		& " WHERE n.CtNodeID=" & pkstr(request.querystring("ctNodeID"),"")
	set RS = Conn.execute(sql)
	session("CtUnitID") = RS("CtUnitID")
	session("ctUnitName") = RS("CtUnitName")
	session("iBaseDSD") = RS("iBaseDSD")
	session("fCtUnitOnly") = RS("fCtUnitOnly")
	session("pvXdmp") = RS("pvXdmp")	
	if isNull(RS("sBaseTableName")) then
		session("sBaseTableName") = "CuDTx" & session("iBaseDSD")
	else
		session("sBaseTableName") = RS("sBaseTableName")
	end if	
	if isNumeric(RS("iBaseDSD")) then
%>
<script language=vbs>
	window.navigate "DsdXMLin.asp?iBaseDSD=<%=RS("iBaseDSD")%>"
</script>
<%  	response.end
	elseif RS("ctUnitKind")="U" then
%>
<script language=vbs>
	window.navigate "CtUnitEdit.asp?CtUnitID=<%=RS("CtUnitID")%>"
</script>
<%  	response.end
	end if 
%>
