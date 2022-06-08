<%@ CodePage = 65001 %><%
Set Conn = Server.CreateObject("ADODB.Connection")
xStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mapPath("/0inDB/") & "\Important.mdb"
response.write xStr & "<HR>"

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open   xStr
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = xStr
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------



	sqlCom = "SELECT * FROM [重要措施]" 
	set RS = conn.execute(sqlCOM)

	response.write "<TABLE BORDER><TR class=eTableLable>"
	for i = 0 to RS.fields.count-1
		response.write "<TH>" & RS.fields(i).name
	next

	while not RS.eof

		response.write "<TR>"
		for i = 0 to RS.fields.count-1
			response.write "<TD>" & RS.fields(i)
		next

		
		RS.moveNext
	wend
	response.write "</TABLE>"

%>
</body>
</html>








