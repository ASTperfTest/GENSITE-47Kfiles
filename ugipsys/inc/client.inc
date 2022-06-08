<%

sub aaa()
	Set Conn = Server.CreateObject("ADODB.Connection")

	'----------HyWeb GIP DB CONNECTION PATCH----------
	'Conn.Open session("ODBCDSN")
	'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
	Conn.ConnectionString = session("ODBCDSN")
	Conn.ConnectionTimeout=0
	Conn.CursorLocation = 3
	Conn.open
	'----------HyWeb GIP DB CONNECTION PATCH----------


	Set KMConn = Server.CreateObject("ADODB.Connection")

	'----------HyWeb GIP DB CONNECTION PATCH----------
	'KMConn.Open session("KMDBConnStr")
	'Set KMConn = Server.CreateObject("HyWebDB3.dbExecute")
	KMConn.ConnectionString = session("KMDBConnStr")
	KMConn.ConnectionTimeout=0
	KMConn.CursorLocation = 3
	KMConn.open
	'----------HyWeb GIP DB CONNECTION PATCH----------

end sub



Set KMConn = Server.CreateObject("ADODB.Connection")
KMConn.ConnectionString = session("KMDBConnStr")
KMConn.ConnectionTimeout=0
KMConn.open

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionString = session("ODBCDSN")
conn.ConnectionTimeout=0
conn.CursorLocation = 3
conn.open



%>
