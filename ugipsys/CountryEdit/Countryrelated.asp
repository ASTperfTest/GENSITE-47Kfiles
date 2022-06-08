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


	country = request("country")
	response.write country & "<BR/>"
	sql = "Select mcode FROM CodeMain " & _
		"where mvalue='"&country&"' AND codeMetaId='country_edit' Order BY mcode"
	response.write sql
	set RS = conn.execute(sql)
	while not RS.eof
		response.write "<row><mCode>" & RS("mcode") & "</mCode><mValue><![CDATA[" & RS("mcode") & "]]></mValue></row>" & vbCRLF
		RS.moveNext
	wend
	
%>
</divList>
