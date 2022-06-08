<%
	Response.Expires = 0

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

	set vo = oxml.selectSingleNode("UseCase/Version[Date='"	_
		& request.queryString("verDate") & "']")	
  if nullText(vo) = "" then
  	response.write "no matched version"
	response.end
  end if

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

	session("verDate") = request.queryString("verDate")
	session("apCode") = request.queryString("apCode")
	set oxsl = server.createObject("microsoft.XMLDOM")
	  oxsl.async = false
	if Instr(session("uGrpID")&",", "HTSD,") > 0 then  
		xv = oxsl.load(server.mappath("/HTSDSystem/ucExpanded.xsl"))
	else
		xv = oxsl.load(server.mappath("/HTSDSystem/uuExpanded.xsl"))
	end if
'	response.write vo.transformNode(oxsl)
	Response.ContentType = "text/HTML" 
	response.write replace(vo.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	response.end



	
%>
