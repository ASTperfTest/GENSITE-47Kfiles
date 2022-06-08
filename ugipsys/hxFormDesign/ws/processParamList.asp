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


	repeatTimes = request("repeatTimes")
	xTabIndex = request("xtIndex")
  	set htPageDom = session("hyXFormSpec")
   	set xdsRoot = htPageDom.selectSingleNode("//dsTable[tableName='formList']")
   	xdsRoot.selectSingleNode("repeatTimes").text = repeatTimes
	set session("hyXFormSpec") = htPageDom
   	set dsRoot = xdsRoot.cloneNode(true)

		response.write "<TABLE border=""1"" summary=""多列式輸入表格"">" _
			& "<CAPTION>(多列式輸入表格，請逐項填寫)</CAPTION><TR>"
		if request("withNum")="1" then	response.write "<TH scope=""col"">項次</TH>"
		for each param in dsRoot.selectNodes("fieldList/field")
			param.selectSingleNode("fieldName").text = param.selectSingleNode("fieldName").text & "_00"
			response.write "<TH scope=""col"">"
			xTag = nulltext(param.selectSingleNode("fieldDesc"))
			if xTag = "" then	xTag = nulltext(param.selectSingleNode("fieldLabel"))
			response.write xTag
			response.write "</TH>"
		next
		response.write "</TR>"

	for xi = 1 to repeatTimes
		xn = right("00" & xi, 2)
		response.write "<TR>"
		if request("withNum")="1" then	response.write "<TD align=""center"">" & xi & "</TD>"
		for each param in dsRoot.selectNodes("fieldList/field")
			param.selectSingleNode("fieldName").text = _
				left(param.selectSingleNode("fieldName").text,len(param.selectSingleNode("fieldName").text)-2) & xn
			response.write "<TD align=""center"">"
	'		response.write param.xml
		   	processParamField param, xTabIndex, false
		   	xTabIndex = xTabIndex + 1
			response.write writeCodeStr
			response.write "</TD>"
		next
		response.write "</TR>"
	
	next

	response.write "</TABLE>"
	
%>]]>
</pFieldHTML>
