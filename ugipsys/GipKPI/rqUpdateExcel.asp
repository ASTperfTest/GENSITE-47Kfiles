<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="績效管理"
HTProgFunc="稽催清單"
HTProgCode="GC1AP9"
HTProgPrefix="kpiQuery" 
        Response.ContentType = "application/vnd.ms-excel"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
dim xcatCount(10)

if fSql="" then
	fSql = "SELECT htx.ctUnitId, htx.ctUnitName, htx.ctunitexpireday " _
		& " , s.userId, u.userName, u.email" _
		& " , (SELECT max(deditDate) FROM CuDtGeneric WHERE ictunit=htx.ctUnitId) AS lastUpdate" _
		& " , n.catName" _
		& " FROM CtUnit AS htx JOIN CatTreeNode AS n ON n.ctUnitId=htx.ctUnitId" _
		& " JOIN CtUserSet AS s ON s.ctNodeId=n.ctNodeId" _
		& " JOIN InfoUser AS u ON u.userId=s.userId" _
		& " WHERE htx.ctunitexpireday > 0" _
		& " AND (SELECT max(deditDate) FROM CuDtGeneric WHERE ictunit=htx.ctUnitId) " _
  		& " < DATEADD(day, -htx.ctunitexpireday, getdate())" _
		& " ORDER BY htx.ctUnitId, n.ctNodeId, s.userId" 
'		& " JOIN InfoUser AS u ON u.userId=s.userId -- AND u.email is not NULL" _
end if
'	response.write fSql & "<HR>"
	set RS = conn.execute(fSql)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body>
        <table border="1"  cellspacing="1" class="bluetable" cellpadding="2">
          <tr  class="lightbluetable">
            <th>單元名稱</th>
            <th>更新頻率</th>
            <th>最後更新</th>
            <th>節點項目</th>
            <th>負 責 人</th>
          </tr>
<%
	xCtUnit = ""
	while not RS.eof
		xDept=RS("CtUnitID")
		if xCtUnit <> RS("CtUnitID") then
			if xCtUnit <> "" then		
				response.write "</td></tr>"
			end if

			xCtUnit = RS("CtUnitID")
%>
            <tr style="background:#ffffff"><td><%=RS("CtUnitName")%></td>
            <td align="center"><%=RS("CtUnitexpireday")%> 天
            <td align="center"><%=RS("lastUpdate")%>
            <td><%=RS("CatName")%>
            <td>
<%
		end if
%>
		‧<A href="rqUpdateMail.asp?user=<%=RS("userID")%>&unit=<%=xCtUnit%>"><%=RS("userName")%></A>
<%
		RS.moveNext
	wend    
 
%>       
		</td>
            </tr>
		</table>

</body>
</html>