﻿<%@ CodePage = 65001 %>
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
<title>選擇連結網頁</title>
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
	fSql = "SELECT icuitem, stitle from CuDtGeneric where ictunit=" & session("CtUnitID") _
		& " AND icuitem<>" & xiCUItem _
		& " ORDER BY icuitem DESC"
	set RSlist = conn.execute(fSql)
%>
 <TABLE width="100%" cellspacing="1" cellpadding="0" class="bg" border=1>
 <tr align="left">
	<td class=eTableLable>連結網頁</td>
 </tr>
<%
	while not RSlist.eof
%>
	<TD class=eTableContent><font size=2>
	<SPAN onclick="ItemSelected('<%=RSlist("iCuItem")%>');" style="cursor:hand;">
<%=RSlist("sTitle")%></SPAN>
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
