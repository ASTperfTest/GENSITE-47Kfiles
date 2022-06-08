<?xml version="1.0"  encoding="utf-8" ?>
<divList>
<%
	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
'Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
conn.ConnectionTimeout=0
conn.CursorLocation = 3
conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	iBaseDSD = request("iBaseDSD")
	sql = "Select CtUnitID,CtUnitName from CtUnit where iBaseDSD='"&iBaseDSD&"' Order by CtUnitID"
	set RS = conn.execute(sql)
	while not RS.eof
		response.write "<row><mCode>" & RS("CtUnitID") & "</mCode><mValue><![CDATA[" & RS("CtUnitName") & "]]></mValue></row>" & vbCRLF
		RS.moveNext
	wend
	
%>
</divList>
