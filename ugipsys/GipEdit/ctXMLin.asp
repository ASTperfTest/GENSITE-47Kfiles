<%@ CodePage = 65001 %>
<% 
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%	

	sql = "SELECT u.*, u.ibaseDsd, b.sbaseTableName, r.pvXdmp, catName FROM CatTreeNode AS n LEFT JOIN CtUnit AS u ON u.ctUnitId=n.ctUnitId" _
		& " Left Join BaseDsd As b ON u.ibaseDsd=b.ibaseDsd" _
		& " LEFT JOIN CatTreeRoot AS r ON r.ctRootId=n.ctRootId" _
		& " WHERE n.ctNodeId=" & pkstr(request.querystring("ctNodeId"),"")
'	response.write sql
'	response.end
	session("ctNodeId") = request("ctNodeId")	
	session("onlyThisNode") = ""
	if checkGIPconfig("onlyThisNode") then
		set RS = conn.execute("SELECT onlyThisNode FROM CatTreeNode WHERE ctNodeId=" & pkstr(request.querystring("ctNodeId"),""))
		session("onlyThisNode") = RS("onlyThisNode")
	end if
	set RS = Conn.execute(sql)
	session("ItemID") = request("ItemID")
	session("ctUnitId") = RS("ctUnitId")
	session("ctUnitName") = RS("CtUnitName")
	session("catName") = RS("catName")
	session("ibaseDsd") = RS("ibaseDsd")
	session("fCtUnitOnly") = RS("fCtUnitOnly")
	session("pvXdmp") = RS("pvXdmp")
	if isNull(RS("sbaseTableName")) then
		session("sbaseTableName") = "CuDTx" & session("ibaseDsd")
	else
		session("sbaseTableName") = RS("sbaseTableName")
	end if
	session("CheckYN") = RS("CheckYN")	
	session("shortLongList") = ""
	if RS("ctUnitKind")="U"  and trim(RS("ibaseDsd") & " ") = "" then
%>
<script language=vbs>
'	window.navigate "CtUnitEdit.asp?ctUnitId=<%=RS("ctUnitId")%>&phase=edit"
</script>
<%  	response.end
	elseif isnull(RS("ctUnitId")) or session("ctUnitId")="" then
'		response.write "Folder Node without data" 
	  	response.end
	else
%>
<script language=vbs>
	window.navigate "DsdXMLin.asp?ibaseDsd=<%=RS("ibaseDsd")%>&ctNodeId=<%=request.querystring("ctNodeId")%>"
</script>
<%  	response.end
	end if 
%>
