<% 
	Response.Expires = 0
	HTProgCap = "討論專區"
	HTProgFunc = "討論專區維護"
	HTProgCode = "BBS010"
	HTProgPrefix = "mSession" 
	response.expires = 0 
%>
<!-- #include virtual = "/inc/server.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
 
	sql = "update talk03 set content = '" & replace(trim(request("content")), "'", "''") & "', updatetime = getdate()"
	conn.execute(sql)
	
	response.redirect "4_1.asp"
%>