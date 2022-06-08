<%

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "Provider=SQLOLEDB;Data Source=127.0.0.1;User ID=hygip;Password=hyweb;Initial Catalog=mGIPcoanew"
%>

<%
	sql = "SELECT * FROM HyftdIndexBuild WHERE CtNodeID = 0"
	
	set rs = conn.execute(sql)
		
	while not rs.eof 
	
		siteId = rs("siteId")
		icuItem = rs("icuItem") 
		baseDsd = rs("baseDsd") 
		ctUnit = rs("ctUnit") 
		ctNodeId = rs("ctNodeId")
		
		sql = "INSERT INTO HyftdIndexDelete VALUES(" & siteId & ", " & icuItem & ", " & _
					baseDsd & ", " & ctUnit & ", " & ctNodeId & ", GETDATE(), 0) "
		'response.write sql & "<br>"
		conn.execute(sql)
		
		rs.movenext		
	wend
	rs.close
	set rs = nothing
	
	sql = "DELETE FROM HyftdIndexBuild WHERE CtNodeID = 0"
	conn.execute(sql)

	
%>