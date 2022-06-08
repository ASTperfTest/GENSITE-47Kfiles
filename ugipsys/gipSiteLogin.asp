<%@ CodePage = 65001 %><%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1
%>
<!--#INCLUDE FILE="inc/dbutil.inc" -->
<%
'	session("GIPODBC")="Provider=SQLOLEDB;Data Source=127.0.0.1;User ID=hyGIP;Password=hyweb;Initial Catalog=mGIPhy"
'	session("ODBCDSN")="Provider=SQLOLEDB;Data Source=127.0.0.1;User ID=hyGIP;Password=hyweb;Initial Catalog=mGIPhy"

Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("GIPODBC")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("GIPODBC")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	sql = "SELECT * FROM GipSites where GipSiteID=" & pkStr(request("siteID"),"")

	set RS = conn.execute(sql)
	if RS.eof then
		response.write "Site not found!!!"
		response.end
	end if

	session("ODBCDSN")=RS("GipSiteDBConn")
	session("mySiteID") = RS("GipSiteID")
'	session("mySiteURL") = "http://hygipsys.hyweb.com.tw"
	session("mySiteURL") = RS("GipSiteURLsys")
	session("mySysSiteURL") = RS("GipSiteURLsys")
	session("myWWWSiteURL") = RS("GipSiteURL")
	session("mySiteName") = RS("GipSiteName")
     session("Public") = "/site/" & RS("GipSiteID") & "/public/"
	session("PageCount") = 10
     Session.Timeout = 60
     		session("pwd") = False

     '---
Set xonn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'xonn.Open session("ODBCDSN")
'Set xonn = Server.CreateObject("HyWebDB3.dbExecute")
xonn.ConnectionString = session("ODBCDSN")
xonn.ConnectionTimeout=0
xonn.CursorLocation = 3
xonn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	session("xRootID") = RS("GipSiteDefaultRoot")
	session("exRootID") = RS("GipSiteInnerRoot")

	sql = "SELECT * FROM CatTreeRoot WHERE vGroup='GY'"
	set RS = xonn.execute(sql)
	if not RS.eof then			session("exRootID") = RS("CtRootID")


	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
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
		top.location.href="default.asp"
<%	end if %>
</script>
</BODY>
</HTML>
<%
	end if
%>
%>









