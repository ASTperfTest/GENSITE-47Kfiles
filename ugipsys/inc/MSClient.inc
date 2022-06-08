<%

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
