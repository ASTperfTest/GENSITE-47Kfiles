<%@CodePage = 65001 %><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #INCLUDE FILE="inc/dbFunc.inc" -->
<%
set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	sqlCom = "SELECT i.*, d.deptName, t.mvalue AS topDataCat " _
		& " FROM InfoUser AS i LEFT JOIN Dept AS d ON i.deptId=d.deptId " _
		& " LEFT JOIN CodeMain AS t ON t.codeMetaId='topDataCat' AND t.mcode=i.tdataCat" _
		& " WHERE userId='" & session("userID") & "'"
	set RS = conn.execute(sqlCom)	
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=session("mySiteName")%></title>
<meta http-equiv="refresh" content="60"/>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/tiltle.css" rel="stylesheet" type="text/css">
</script>
</head>

<body>
<input TYPE=hidden name=mySiteID value="<%=session("mySiteID")%>">
<div id="Title">
	<div><img src="images/title.gif" alt="HyGIP 政府資訊入口網站 管理系統"></div>
	<div id="SiteName"><%=session("mySiteName")%></div>
	<div id="LogInfo">
	<%=RS("userName")%> (<%=RS("userID")%>,<%=RS("deptName")%>-<%=RS("topDataCat")%>) 
	  上次於 <%=d7date(session("lastVisit"))%> 由<%=session("lastIP")%>第<%=session("VisitCount")%>次登入
	</div>
	<a href="logout.asp?siteID=<%=session("mySiteID")%>" class="Logout">登出</a>
</div>
</body>
</html>
<%
	Conn.close()
	set Conn = nothing
%>


<script language=VBS>  
sub sessionOutReLoad()
	top.location.href="gipSiteEntry.asp?siteID="&document.all("mySiteID").value
end sub
</script>
