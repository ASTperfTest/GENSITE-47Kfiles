<% Response.Expires = 0
HTProgCap="活動管理"
HTProgFunc="報名學員清單"
HTProgCode="PA005"
HTProgPrefix="psEnroll" %>
<% response.expires = 0 %>
<!--#Include file = "../inc/server.inc" -->
<%
	sql = "SELECT * FROM paSession WHERE paSID=" & session("paSID")
	set RS = conn.execute(sql)
	pLimit = RS("pLimit")
	pBackup = RS("pBackup")
	
	sql = "UPDATE paEnroll set status = 'N'" _
		& " WHERE pasID=" & session("paSID")
	conn.execute sql
		
			
	sql = "UPDATE paEnroll set status = 'B'" _
		& " WHERE pasID=" & session("paSID") & " AND psnID IN" _
		& " (SELECT top " & pLimit+pBackup & " psnID from paEnroll where pasid=" & session("paSID") _
		& " ORDER BY ckValue DESC)"
	conn.execute sql

	sql = "UPDATE paEnroll set status = 'Y'" _
		& " WHERE pasID=" & session("paSID") & " AND psnID IN" _
		& " (SELECT top " & pLimit & " psnID from paEnroll where pasid=" & session("paSID") _
		& " ORDER BY ckValue DESC)"
	conn.execute sql

%>
<Html>
<body>
<script language=vbs>
	window.navigate "psEnrollList.asp"
</script>	
</body>
</html>                                 
