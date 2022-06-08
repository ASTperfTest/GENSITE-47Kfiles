<%
	response.expires = 0
	apCode = session("apCode")
	verDate = session("verDate")
	xid = request.queryString("id")
	xmlFile = "evx_" & request.queryString("xmlFile")
	
	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true
	LoadXML = server.mappath("/HTSDSystem/UseCase/" & apCode & ".xml")
	xv = oxml.load(LoadXML)
  if oxml.parseError.reason <> "" then
    Response.write loadXML & "<BR>"
    Response.Write("XML parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if

	set vo = oxml.selectSingleNode("UseCase/Version[Date='"	& verDate & "']")
  if nullText(vo) = "" then
  	response.write "version not exist"
	response.end
  end if

	set nxml = server.createObject("microsoft.XMLDOM")
	nxml.async = false
	nxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/HTSDSystem/contract0.xml")
	xv = nxml.load(LoadXML)
  if nxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  nxml.parseError.line)
    Response.Write("<BR>Reason: " &  nxml.parseError.reason)
    Response.End()
  end if

  	set eo = vo.selectSingleNode("//TypicalCourseOfEvents/Event[@id='" & xid & "']")
  	eo.setAttribute "ContractXML", xmlFile&".xml"


'  response.write oxml.selectSingleNode("//frame[@name='topFrame']/@src").text

	nxml.selectSingleNode("OperationContract/Code").text = apCode
	nxml.selectSingleNode("OperationContract/APCat").text = nullText(oxml.selectSingleNode("UseCase/APCat"))
	nxml.selectSingleNode("OperationContract/Version/Date").text = date()
	nxml.selectSingleNode("OperationContract/Version/Author").text = session("userID")
	nxml.selectSingleNode("OperationContract/Version/FromEvent").setAttribute "type" _
		, nullText(eo.selectSingleNode("@type"))
	nxml.selectSingleNode("OperationContract/Version/FromEvent").setAttribute "id", xid

	nxml.selectSingleNode("OperationContract/Version/ContractSpec/Name").text = xmlFile
	nxml.selectSingleNode("OperationContract/Version/ContractSpec/Type").text = nullText(eo.selectSingleNode("@type"))
	nxml.selectSingleNode("OperationContract/Version/ContractSpec/ResponsibilityList/Responsibility").text _
		= nullText(eo)


'  response.write oxml.documentElement.xml
	nxml.save(Server.MapPath("/HTSDSystem/UseCase/" & xmlFile & ".xml"))

  
	oxml.save Server.MapPath("/HTSDSystem/UseCase/" & apCode & ".xml")
  	
	response.redirect "ucExpanded.asp?apCode=" & apCode & "&verDate=" & verDate

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

%>
