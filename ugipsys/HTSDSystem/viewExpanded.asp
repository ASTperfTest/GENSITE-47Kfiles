<%
	Response.Expires = 0

	apCode = session("apCode")

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true
	LoadXML = server.mappath("/HTSDSystem/UseCase/" & apCode & ".xml")
'	response.write loadxml & "<HR>"
'	response.end
	xv = oxml.load(LoadXML)
	set vo = oxml.selectSingleNode("UseCase/Version[Date='"	_
		& session("verDate") & "']")	
'response.write vo.text
'response.end		
  if nullText(vo) = "" then
  	response.write "no matched version"
	response.end
  end if	

  if oxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

	set oxsl = server.createObject("microsoft.XMLDOM")
	oxsl.async = false
	xAPPattern = nullText(vo.selectSingleNode("ExpandedSpec/APPattern"))
	select case xAPPattern
		case "Query"
			xv = oxsl.load(server.mappath("/HTSDSystem/viewExpanded_Query.xsl"))	
		case else
			response.write "Bye"
	end select
'	response.write vo.transformNode(oxsl)
	Response.ContentType = "text/HTML" 
	response.write replace(oxml.selectSingleNode("//Version").transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	response.end
%>
