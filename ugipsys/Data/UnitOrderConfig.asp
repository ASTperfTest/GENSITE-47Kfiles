<%@ CodePage = 65001 %>
<!--#Include file = "index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<%
 xType = Request.QueryString("xType")
 CatSQL = Request.QueryString("CatSQL")
 If xType = "UnitOrderBy" Then
	UnitID = Request.Form("UnitID")
	UnitIDArrar = Split(UnitID,",")
	 NowShowOrder = 0
	  For xno = 0 to UBound(UnitIDArrar)
	   NowShowOrder = NowShowOrder + 1
		sql2 = "Update DataUnit set ShowOrder =" & NowShowOrder & " Where UnitID ="& UnitIDArrar(xno)
		set rs2 = conn.execute(sql2)
	  Next
	  msg = "排序完成！"
 End IF
%>
<script language=VBScript>
  alert("<%=msg%>")
  document.location.href="BulletinList.asp?Language=<%=Language%>&DataType=<%=DataType%>&CatSQL=<%=CatSQL%>"
</script>
