<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="資料附件"
HTProgFunc="條列"
HTProgCode="GC1AP1"
HTProgPrefix="CuAttach" %>
<% response.expires = 0 %>
<!--#Include file = "../inc/server.inc" -->
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="../inc/setstyle.css">
<title></title>
</head>
<!--#INCLUDE FILE="../inc/dbFunc.inc" -->
<%
Dim mpKey, dpKey
Dim RSmaster, RSlist
	xiCUItem = request.queryString("iCUItem")
%>
<body>
<SCRIPT>
function ItemSelected(c) {
window.returnValue = c;
window.close();
}
</SCRIPT>
<%
	fSql = "SELECT dhtx.*, r.sTitle" _
		& " FROM CuDTPage AS dhtx" _
		& " LEFT JOIN CuDTGeneric AS r ON r.iCuItem=dhtx.nPageID" _
		& " WHERE 1=1" _
		& " AND dhtx.xiCUItem=" & pkStr(xiCUItem,"") _
		& " ORDER BY dhtx.listSeq"
	set RSlist = conn.execute(fSql)
%>
 <TABLE width="100%" cellspacing="1" cellpadding="0" class="bg">
 <tr align="left">
	<td class=eTableLable>連結名稱</td>
	<td class=eTableLable>說明</td>
	<td class=eTableLable>連結頁</td>
 </tr>
<%
	while not RSlist.eof
%>
	<TD class=eTableContent><font size=2>
	<SPAN onclick="ItemSelected('<%=RSlist("nPageID")%>');" style="cursor:hand;">
<%=RSlist("aTitle")%></SPAN>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("aDesc")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("sTitle")%>
</font></td>
    </tr>
    <%
         RSlist.moveNext
     wend
   %>
    </TABLE>
    </CENTER>
       <!-- 程式結束 ---------------------------------->  
     </form>  
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 
</body>
</html>                                 
