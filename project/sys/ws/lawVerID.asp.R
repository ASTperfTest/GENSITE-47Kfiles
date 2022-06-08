<?xml version="1.0" encoding="utf-8"?>
<divList>
<%

	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open session("ODBCDSN")

	lawID = request("lawID")
	sql = "Select iCuItem,sTitle from CuDTGeneric where topCat='"&lawID&"' and ictunit=1988"
	set RS = conn.execute(sql)
	while not RS.eof
		response.write "<row><mCode>" & RS("iCuItem") & "</mCode><mValue><![CDATA[" & RS("sTitle") & "]]></mValue></row>" & vbCRLF
		RS.moveNext
	wend
	
%>
</divList>
