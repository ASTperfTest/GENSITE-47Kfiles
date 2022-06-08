<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1

'	ODBCDSN="Provider=SQLOLEDB;Data Source=192.168.0.21;User ID=hyGIP;Password=hyweb;Initial Catalog=GIPdoca"
	DBConnStr="Provider=SQLOLEDB;server=127.0.0.1;User ID=hyGIP;Password=hyweb;Database=GIPmofDB"
'	response.write "<a>" & session("ODBCDSN") & "</a></divList>"
'	response.end
	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open DBConnStr
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = DBConnStr
'----------HyWeb GIP DB CONNECTION PATCH----------


	fDate = request("date")
	if fDate="" then
		sql = "SELECT max(fidxDate) FROM FinanceIndexDaily"
		set RS = conn.execute(sql)
		fDate = RS(0)
	end if

  cvbCRLF = vbCRLF
  cTabchar = chr(9)
'  cvbCRLF = ""
'  cTabchar = ""

	xmlStr = "<?xml version=""1.0"" encoding=""utf-8"" ?>" & cvbCRLF
'	xmlStr = ""
	xmlStr = xmlStr & "<fidxList>" & cvbCRLF
	xmlStr = xmlStr & ncTabchar(1) & "<fidxDate>" & fDate & "</fidxDate>" & cvbCRLF

	sql = "SELECT h.*, c.fidxNameE FROM FinanceIndexDaily AS h JOIN FinanceIndex AS c ON h.fidxID=c.fidxID" _
		& " WHERE fidxDate=N'" & fDate & "'"
	set RS = conn.execute(sql)
	while not RS.eof
		xmlStr = xmlStr & ncTabchar(2) & "<fidxItem><eCode>" & RS("fidxNameE") & "</eCode><fValue>" & RS("fidxPrice") & "</fValue></fidxItem>" & vbCRLF
		RS.moveNext
	wend
	xmlStr = xmlStr & "</fidxList>" & cvbCRLF
'	debugPrint xmlStr

  set oXmlReg = Server.CreateObject("MICROSOFT.XMLDOM")
  oXmlReg.async = false
'  debugPrint xmlStr
'  response.end
  oXmlReg.loadXML xmlStr

  postURL = "http://smof.cw/ws/fidxRecv.asp"
  postURL = "http://wadmin.mof.gov.tw/ws/fidxRecv.asp"
'  debugPrint postURL & "<HR>"

	set xmlHTTP = Server.CreateObject("MSXML2.serverXMLHTTP")
	set rXmlObj = Server.CreateObject("MICROSOFT.XMLDOM")
	rXmlObj.async = false
	xmlHTTP.open "POST", postURL, false
	xmlHTTp.send oXmlReg.xml
	rv = rXmlObj.load(xmlHTTP.responseXML)
	if not rv then
		debugPrint  xmlHTTP.responseTEXT & "<HR>"
		response.write rv & "<HR>"
		response.end
	end if
'	response.write "obj.xml==><BR>" & rXmlObj.xml & "<HR>"
	debugPrint "<?xml version=""1.0"" encoding=""utf-8"" ?>" & cvbCRLF
	debugPrint rXmlObj.documentElement.xml
	
	
function ncTabChar(n)
	if cTabchar = "" then
		ncTabChar = ""
	else
		ncTabChar = String(n,cTabchar)
	end if
end function

sub debugPrint(xstr)
	response.write xstr
end sub

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
