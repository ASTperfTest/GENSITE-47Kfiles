<%
	Response.Expires = 0

	evxFile = request.queryString("xmlFile")

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true
	LoadXML = server.mappath("/HTSDSystem/UseCase/" & evxFile)
'	response.write loadxml & "<HR>"
'	response.end
	xv = oxml.load(LoadXML)
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
'response.write "Hello1[]" & session("uGrpID") & "<br>"	  
	if Instr(session("uGrpID")&",", "HTSD,") > 0 then  
'response.write "Hello2<br>"		
		xv = oxsl.load(server.mappath("/HTSDSystem/viewContract.xsl"))
	else
'response.write "Hell03<br>"		
		xv = oxsl.load(server.mappath("/HTSDSystem/uviewContract.xsl"))
	end if
'	response.write vo.transformNode(oxsl)
	Response.ContentType = "text/HTML" 
	response.write replace(oxml.selectSingleNode("//Version").transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	response.end
%>
