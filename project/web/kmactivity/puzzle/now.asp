<%
	set conn = server.createobject("adodb.connection")
	Conn.ConnectionString = application("ConnStrPuzzle")
	Conn.ConnectionTimeout=0
	Conn.CursorLocation = 3
	Conn.open
	
	
	set rs = conn.execute ("select * from sso")
	
	response.write rs(0)
%>