<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap="局長/民意 信箱"
HTProgFunc="局長信箱維護"
HTProgCode="BM010"
HTProgPrefix="mSession"
response.expires = 0
%>
<!-- #include virtual = "/inc/server.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	
	sql = "" & _
		" update mailbox set " & _
		" email = '" & replace(trim(request("email")), "'", "''") & "' " & _
		" where mailord = '" & replace(request("mailord"), "'", "''") & "'"
	'response.write sql
	conn.execute(sql)
	response.write "<script language='javascript'>alert('Down!');location.replace('index.asp');</script>"
%>