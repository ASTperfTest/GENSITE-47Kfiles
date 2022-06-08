<?xml version="1.0"  encoding="utf-8" ?>
<!--#include FILE = "htUIGen.inc" -->
<pFieldHTML><![CDATA[
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


  	set htPageDom = session("hyXFormSpec")
   	set dsRoot = htPageDom.selectSingleNode("DataSchemaDef") 	

	for each param in htPageDom.selectNodes("//field[fieldName='" & request("fname") & "']")
'		response.write param.xml
	   	processParamField param	, request("xtIndex"), true    
		response.write writeCodeStr
	next
	
%>]]>
</pFieldHTML>
