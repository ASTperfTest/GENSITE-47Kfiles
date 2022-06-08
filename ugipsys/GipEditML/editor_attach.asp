<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="資料附件"
HTProgFunc="條列"
HTProgCode="GC1AP1"
HTProgPrefix="CuAttach" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/server.inc" -->
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>選取連結附件</title>
</head>
<!--#include virtual = "/inc/dbFunc.inc" -->
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
	fSql = "SELECT dhtx.*, i.oldFileName " _
		& " FROM CuDtattach AS dhtx" _
		& " LEFT JOIN ImageFile AS i ON i.newFileName=dhtx.nfileName" _
		& " WHERE 1=1" _
		& " AND dhtx.xiCuItem=" & pkStr(xiCUItem,"") _
		& " ORDER BY dhtx.listSeq"
	set RSlist = conn.execute(fSql)
%>
 <TABLE width="100%" cellspacing="1" cellpadding="0" class="bg">
 <tr align="left">
	<td class=eTableLable>附件名稱</td>
	<td class=eTableLable>原檔名</td>
	<td class=eTableLable>上傳人</td>
	<td class=eTableLable>上傳日</td>
 </tr>
<%
	while not RSlist.eof
		dpKey = ""
		dpKey = dpKey & "&ixCuAttach=" & RSlist("ixCuAttach")
		if dpKey<>"" then  dpKey = mid(dpKey,2)
%>
	<TD class=eTableContent><font size=2>
	<SPAN onclick="ItemSelected('<%=RSlist("NFileName")%>');" style="cursor:hand;">
<%=RSlist("aTitle")%></SPAN>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("OldFileName")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("aEditor")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("aEditDate")%>
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
