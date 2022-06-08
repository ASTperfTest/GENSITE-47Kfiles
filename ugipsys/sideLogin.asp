<%@ CodePage = 65001 %><% Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1 %>
<!-- #INCLUDE FILE="inc/dbutil.inc" -->
<%

	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
'Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
conn.ConnectionTimeout=0
conn.CursorLocation = 3
conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

	sqlCom = "SELECT i.*, d.tDataCat " _
		& " FROM InfoUser AS i LEFT JOIN dept AS d ON i.deptID=d.deptID " _
		& " WHERE userID=" & pkStr(request("userName"),"") _
		& " AND password=" & pkStr(request("passWord"),"")
'	response.write sqlCom
'	response.end ' or 'a'='a
	set RS = conn.execute(sqlCom)
	if RS.eof then
%>
<HTML>
<BODY>
<script language=vbs>
	msgBox "帳號/密碼 不合！請重新簽入！"
	window.history.back
</script>
<%
	else
		session("lastVisit") = RS("LastVisit")
		session("lastIP") = RS("lastIP")
		session("VisitCount") = RS("VisitCount")
		LastVisit = Date()
		VisitCount = RS("VisitCount") + 1
		SQL = "Update InfoUser Set LastVisit = '"& LastVisit &"', VisitCount = "& VisitCount _
			& ", lastIP='" & request.serverVariables("REMOTE_ADDR") &"' Where UserID = '"& RS("UserID") &"'"
		conn.execute(SQL)
		session("pwd") = True
		session("userID") = RS("userID")
		session("userName") = RS("userName")
		session("uGrpID") = RS("uGrpID")
		session("deptID") = RS("deptID")
		session("tDataCat") = RS("tDataCat")
		xPath = "/public/data/" & trim(RS("uploadPath")) 
		if right(xPath,1)<>"/" then	xPath = xPath & "/"
		session("uploadPath") = xPath
		session("Email") = RS("Email")
%>
<script language=vbs>
<% 	if Instr(session("uGrpID")&",", "HTSD,") > 0 then %>
		top.location.href="POPDown.htm"
<%	else %>
		top.location.href="/gipSiteEntry.asp?siteID=hygip"
<%	end if %>
</script>
</BODY>
</HTML>
<%
	end if
%>
