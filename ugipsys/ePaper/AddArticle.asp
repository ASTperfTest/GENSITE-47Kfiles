<%
	Response.Expires = 0
	Response.Expiresabsolute = Now() - 1 
	Response.AddHeader "pragma","no-cache" 
	Response.AddHeader "cache-control","private" 
	Response.CacheControl = "no-cache"
	
	On Error resume Next

	Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	Conn.Open session("ODBCDSN")
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
	For i = 0 To UBound(checkarr)	
		arr = empty
		arr = split(checkarr(i), "-")
		sql = "INSERT INTO EpPubArticle VALUES(" & epubid & ", " & arr(0) & ", " & arr(1) & ", '" & arr(2) & "'," & 0 & ")"
		Conn.Execute(sql)						
	Next

	For i = 0 To UBound(uncheckarr)	
		arr = empty
		arr = split(uncheckarr(i), "-")
		sql = "DELETE FROM EpPubArticle WHERE EpubId = " & epubid & " AND ArticleId = " & arr(0) & " AND CtRootId = " & arr(1) & " AND CategoryId = '" & arr(2) & "'"
		Conn.Execute(sql)						
	Next	
	
''	Conn.Close
		Response.Write 1

%>
