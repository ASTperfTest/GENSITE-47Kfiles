<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1
%>
<!--#INCLUDE FILE="inc/dbutil.inc" -->
<%
'	session("GIPODBC")="Provider=SQLOLEDB;Data Source=127.0.0.1;User ID=hygip;Password=hyweb;Initial Catalog=nGiphy"
'	session("ODBCDSN")="Provider=SQLOLEDB;Data Source=127.0.0.1;User ID=hygip;Password=hyweb;Initial Catalog=nGiphy"

	Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	Conn.Open session("GIPODBC")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("GIPODBC")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	sql = "SELECT * FROM GipSites where gipSiteId=" & pkStr(request("siteID"),"")
	
	'response.write session("GIPODBC")

	set RS = conn.execute(sql)
	if RS.eof then
		response.write "Site not found!!!"
		response.end
	end if

	'931116 Start Add By Jason
	'For DGBAS SSO
	session("DGBAS_uid") = request("uid")
	'931116 End Add By Jason

	session("ODBCDSN")=RS("GipSiteDBConn")
	session("mySiteID") = RS("GipSiteID")
	session("mySiteURL") = RS("GipSiteURLsys")
	session("mySysSiteURL") = RS("GipSiteURLsys")
	session("myWWWSiteURL") = RS("GipSiteURL")
	session("mySiteName") = RS("GipSiteName")
    session("Public") = "/site/" & RS("GipSiteID") & "/public/"
    session("MMOPublic") = "/site/" & RS("GipSiteID") & "/public/"	 
	session("PageCount") = 10
    Session.Timeout = 60
    session("pwd") = False

     '---
	Set xonn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	xonn.Open session("ODBCDSN")
'Set xonn = Server.CreateObject("HyWebDB3.dbExecute")
xonn.ConnectionString = session("ODBCDSN")
xonn.ConnectionTimeout=0
xonn.CursorLocation = 3
xonn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	session("xRootID") = RS("GipSiteDefaultRoot")
	session("exRootID") = RS("GipSiteInnerRoot")

	sql = "SELECT * FROM CatTreeRoot WHERE vgroup='GY'"
	set RS = xonn.execute(sql)
	if not RS.eof then			session("exRootID") = RS("CtRootID")

	'931116 Start Add By Jason
	'For DGBAS SSO
	'當有傳入使用者帳號，直接去找資料庫，跳過輸入帳號密碼畫面
	if session("DGBAS_uid")<>"" then
		response.redirect "checkLogin_rdecthinktank.asp"
	else
		response.redirect "default_rdecthinktank.asp"
	end if
	'931116 End Add By Jason
	
%>
