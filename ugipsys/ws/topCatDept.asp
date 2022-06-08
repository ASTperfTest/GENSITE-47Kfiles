<?xml version="1.0"  encoding="utf-8" ?>
<divList>
<%
'	session("ODBCDSN")="driver={SQL Server};server=203.95.187.76;UID=hometown;PWD=2986648;DATABASE=db921"
	session("ODBCDSN")="Provider=SQLOLEDB;Data Source=10.10.5.128;User ID=hyGIP;Password=hyweb;Initial Catalog=GIPmof"
	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------


	catCode = request("catCode")
	sql = "SELECT deptID, deptName, xURL FROM dept AS d JOIN CuDTGeneric AS g" _
		& " ON d.giCuItem=g.iCuItem WHERE xURL<>'' AND tDataCat LIKE N'%" & catCode & "%'"
	set RS = conn.execute(sql)
	while not RS.eof
		response.write "<row><mCode>" & RS("xURL") & "</mCode><mValue>" & RS("deptName") & "</mValue></row>" & vbCRLF
		RS.moveNext
	wend
	
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
