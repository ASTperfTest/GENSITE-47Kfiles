<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GC1AP1" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->

<%
CtNodeID = request("CtNodeID")
if CtNodeID="" then%>
<script language=VBS>
	alert "未選擇任何一個主題單元!"
	window.close
</script>
<%else

CtNodeIDStr=""
CatNameStr=""
MyArray = Split(CtNodeID,",")
for i=0 to UBound(MyArray)
	xPos=instr(MyArray(i),"|||")
	if xPos<>0 then
		CtNodeIDStr=CtNodeIDStr+Left(MyArray(i),xPos-1)+","
		CatNameStr=CatNameStr+mid(MyArray(i),xPos+3)+","
	end if
	response.write MyArray(i) & "<br>"
next
	CtNodeIDStr=Left(CtNodeIDStr,Len(CtNodeIDStr)-1)
	CatNameStr=Left(CatNameStr,Len(CatNameStr)-1)
'response.write CtNodeIDStr & "<br>"
'response.write CatNameStr & "<br>"
%>
<script language=VBS>
	xCtNodeIDStr = "<%=CtNodeIDStr%>"
	xCatNameStr = "<%=CatNameStr%>"
	window.opener.CtUnitID xCtNodeIDStr,xCatNameStr
	window.close
</script>
<%end if%>