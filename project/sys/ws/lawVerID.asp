<?xml version="1.0" encoding="utf-8"?>
<divList>
<%

	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
'Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------
conn.ConnectionTimeout=0
conn.CursorLocation = 3
conn.open

	lawID = request("lawID")
	sql = "Select iCuItem,sTitle from CuDTGeneric where topCat='"&lawID&"' and ictunit=1988"
	set RS = conn.execute(sql)
	while not RS.eof
		response.write "<row><mCode>" & RS("iCuItem") & "</mCode><mValue><![CDATA[" & RS("sTitle") & "]]></mValue></row>" & vbCRLF
		RS.moveNext
	wend
	
%>
</divList>
