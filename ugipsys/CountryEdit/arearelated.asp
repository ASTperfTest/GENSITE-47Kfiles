<% Response.ContentType = "text/xml" %>
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


	area = request("area")
	response.write area & "<BR/>"
	sql = "Select mcode, mvalue from CodeMain " & _
		"where mref='"&area&"' Order by mcode"
	response.write sql
	set RS = conn.execute(sql)
	while not RS.eof
		response.write "<row><mCode>" & RS("mvalue") & "</mCode><mValue><![CDATA[" & RS("mvalue") & "]]></mValue></row>" & vbCRLF
		RS.moveNext
	wend
	
%>
</divList>
