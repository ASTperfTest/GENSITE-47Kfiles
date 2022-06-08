<%
	Response.Expires = 0
	Response.Expiresabsolute = Now() - 1 
	Response.AddHeader "pragma","no-cache" 
	Response.AddHeader "cache-control","private" 
	Response.CacheControl = "no-cache"
	
	

	check = request("check")
	session("jigcheck") = check
	Response.Write "1"
	response.end
	
	Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	'Conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

	
	epubid = request("epubid")
	check = request("check")
	uncheck = request("uncheck")
	ctrootid = request("CtRootId")
	
	checkarr = split(check, ";")
	uncheckarr = split(uncheck, ";")
	
	

	
							

	
	

%>
