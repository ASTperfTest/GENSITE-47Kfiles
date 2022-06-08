<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open session("ODBCDSN")
xRootID=session("xRootID")
%>
