<% 
	Response.Expires = 0
	HTProgCap = "�Q�ױM��"
	HTProgFunc = "�Q�ױM�Ϻ��@"
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