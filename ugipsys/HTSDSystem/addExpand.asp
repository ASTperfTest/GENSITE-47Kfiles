<%
	apCode = request.queryString("apCode")

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true
	LoadXML = server.mappath("/HTSDSystem/UseCase/" & apCode & ".xml")
	xv = oxml.load(LoadXML)
  if oxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if

	set vo = oxml.selectSingleNode("UseCase/Version[(Date='"	& date() & "') and (Kind='Expanded')]")
  if nullText(vo) <> "" then
  	response.write "version already exist"
	response.end
  end if

	set fxml = server.createObject("microsoft.XMLDOM")
	fxml.async = false
	fxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/HTSDSystem/expanded0.xml")
	xv = fxml.load(LoadXML)
  if fxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  fxml.parseError.line)
    Response.Write("<BR>Reason: " &  fxml.parseError.reason)
    Response.End()
  end if

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

	set xFieldNode = fxml.selectSingleNode("ExpandedUseCase/Version").cloneNode(true)
	
	xFieldNode.selectSingleNode("Date").text = date()
	xFieldNode.selectSingleNode("Author").text = session("userID")
	
	oxml.selectSingleNode("UseCase").insertBefore xFieldNode, oxml.selectSingleNode("UseCase/Version")
'	response.write oxml.xml
	oxml.save Server.MapPath("/HTSDSystem/UseCase/" & apCode & ".xml")
	response.redirect "ucVersion.asp?apCode=" & apCode



	
%>
