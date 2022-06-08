<?xml version="1.0"  encoding="utf-8" ?>
<divList>
<%
'	session("ODBCDSN")="driver={SQL Server};server=203.95.187.76;UID=hometown;PWD=2986648;DATABASE=db921"
	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------


	TuitionItemCode = request("TuitionItem")
	sql = "SELECT htx.* FROM AcctItemCode AS htx" _
		& " WHERE htx.accID=" & pkStr(request("accID"),"")
	set RS = conn.execute(sql)
	if not RS.eof then
		response.write "<row><mCode>" & RS("AccID") & "</mCode><mValue>" _
			& RS("AccIDName") & "</mValue><BgtCreatMode>" _
			& RS("BgtCreatMode") & "</BgtCreatMode><BgtCalcField1>" _
			& RS("BgtCalcField1") & "</BgtCalcField1><BgtCalcField2>" _
			& RS("BgtCalcField2") & "</BgtCalcField2><BgtCalcField3>" _
			& RS("BgtCalcField3") & "</BgtCalcField3><BgtCalcField4>" _
			& RS("BgtCalcField4") & "</BgtCalcField4></row>" & vbCRLF
		RS.moveNext
	end if
	
FUNCTION pkStr (s, endchar)
  if s="" then
	pkStr = "null" & endchar
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	pkStr="'" & s & "'" & endchar
  end if
END FUNCTION

sub retrieveByXML
	set rptXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM")
	rptXmlDoc.async = false
	rptXmlDoc.setProperty("ServerHTTPRequest") = true	
	
	LoadXML = server.MapPath(".") & "\dsd.xml"
'	debugPrint LoadXML & "<HR>"
	xv = rptXmlDoc.load(LoadXML)
	
  if rptXmlDoc.parseError.reason <> "" then 
    Response.Write("XML parseError on line " &  rptXmlDoc.parseError.line)
    Response.Write("<BR>Reason: " &  rptXmlDoc.parseError.reason)
    Response.End()
  end if

  	set optionList = rptXmlDoc.selectNodes("//dsTable[tableName='CodeMain']/instance/row[codeMetaID='"&cityCode&"']")
  	for each opItem in optionList
'  		response.write "<divItem><mCode>" & opItem.selectSingleNode("mCode")
		response.write opItem.xml
  	next
'	response.ContentType = "text/xml"
'	XMlDoc2.save(Response)	
'	response.write 	rptXmlDoc.transformNode(xslDom)
'	response.end
end sub
%>
</divList>
