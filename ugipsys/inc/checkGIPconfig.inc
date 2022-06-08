<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

function checkGIPconfig(funcName)
    dim htPageDomCheck
	set htPageDomCheck = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomCheck.async = false
	htPageDomCheck.setProperty("ServerHTTPRequest") = true

	GIPconfigLoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysPara.xml"
	xv = htPageDomCheck.load(GIPconfigLoadXML)
	if htPageDomCheck.parseError.reason <> "" then
		Response.Write("sysPara.xml parseError on line " &  htPageDomCheck.parseError.line)
		Response.Write("<BR>Reasonxx: " &  htPageDomCheck.parseError.reason)
		Response.End()
	end if

  	checkGIPconfig = false
  	if UCASE(nullText(htPageDomCheck.selectSingleNode("SystemParameter/GIPconfig/" &funcName)))="Y" then	_
		checkGIPconfig = true
end function


'// get setting item text from sysPara.xml
'// modify date: 2006/01/06
function getGIPconfigText(funcName)
    dim rtnVal          '// return value
    dim htPageDomCheck
    dim LoadXMLCheck
    dim xvCheck
	rtnVal = ""
	
	set htPageDomCheck = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomCheck.async = false
	htPageDomCheck.setProperty("ServerHTTPRequest") = true

	LoadXMLCheck = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysPara.xml"
	xvCheck = htPageDomCheck.load(LoadXMLCheck)
	if htPageDomCheck.parseError.reason <> "" then
		Response.Write("sysPara.xml parseError on line " &  htPageDomCheck.parseError.line)
		Response.Write("<BR>Reasonxx: " &  htPageDomCheck.parseError.reason)
		Response.End()
	end if

  	rtnVal = nullText(htPageDomCheck.selectSingleNode("SystemParameter/GIPconfig/" &funcName))
  	getGIPconfigText = rtnVal
end function

%>
