<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function


function checkGIPApConfig(funcName)
    dim htPageDomCheck
    dim LoadXMLCheck
    
	set htPageDomCheck = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomCheck.async = false
	htPageDomCheck.setProperty("ServerHTTPRequest") = true

    LoadXMLCheck = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysApPara.xml"
	xv = htPageDomCheck.load(LoadXMLCheck)
	if htPageDomCheck.parseError.reason <> "" then
		Response.Write("sysApPara.xml parseError on line " &  htPageDomCheck.parseError.line)
		'Response.Write("<BR>LoadXML: " &  LoadXML)
		Response.Write("<BR>Reasonxx: " &  htPageDomCheck.parseError.reason)
		Response.End()
	end if

  	checkGIPApConfig = false
  	if UCASE(nullText(htPageDomCheck.selectSingleNode("SystemParameter/GIPconfig/" &funcName)))="Y" then
		checkGIPApConfig = true
    end if
end function



'// get setting item text to sysApPara.xml
'// modify date: 2006/09/20
function getGIPApConfigText(byVal funcName)
    dim htPageDomCheck
    dim LoadXMLCheck
    
	set htPageDomCheck = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomCheck.async = false
	htPageDomCheck.setProperty("ServerHTTPRequest") = true
	if saveText <> "1" or saveText <> "Y" then
	    saveText = "N"
	end if

    LoadXMLCheck = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysApPara.xml"
	xv = htPageDomCheck.load(LoadXMLCheck)
	if htPageDomCheck.parseError.reason <> "" then
		Response.Write("sysApPara.xml parseError on line " &  htPageDomCheck.parseError.line)
		'Response.Write("<BR>LoadXML: " &  LoadXML)
		Response.Write("<BR>Reasonxx: " &  htPageDomCheck.parseError.reason)
		Response.End()
	end if

  	rtnVal = nullText(htPageDomCheck.selectSingleNode("SystemParameter/GIPconfig/" &funcName))
  	getGIPApConfigText = rtnVal
end function


'// save setting item text from sysApPara.xml
'// modify date: 2006/09/20
function saveGIPAPconfig(funcName, byVal saveText)
    dim rtnVal          '// return value
    dim htPageDomCheck
    dim LoadXMLCheck
    dim xvCheck
	rtnVal = ""
	
	set htPageDomCheck = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomCheck.async = false
	htPageDomCheck.setProperty("ServerHTTPRequest") = true

	LoadXMLCheck = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysApPara.xml"
	xvCheck = htPageDomCheck.load(LoadXMLCheck)
	if htPageDomCheck.parseError.reason <> "" then
		Response.Write("sysApPara.xml parseError on line " &  htPageDomCheck.parseError.line)
		'Response.Write("<BR>LoadXML: " &  LoadXML)
		Response.Write("<BR>Reasonxx: " &  htPageDomCheck.parseError.reason)
		Response.End()
	end if

  	htPageDomCheck.selectSingleNode("SystemParameter/GIPconfig/" &funcName).text = saveText
  	rtnVal = htPageDomCheck.save(LoadXMLCheck)
  	saveGIPApConfig = rtnVal
end function


'// get setting item text from sysApPara.xml
'// modify date: 2006/01/06
function saveGIPAPconfigText(funcName, saveText)
    dim rtnVal          '// return value
    dim htPageDomCheck
    dim LoadXMLCheck
    dim xvCheck
	rtnVal = ""
	
	set htPageDomCheck = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomCheck.async = false
	htPageDomCheck.setProperty("ServerHTTPRequest") = true

	LoadXMLCheck = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysApPara.xml"
	xvCheck = htPageDomCheck.load(LoadXMLCheck)
	if htPageDomCheck.parseError.reason <> "" then
		Response.Write("sysApPara.xml parseError on line " &  htPageDomCheck.parseError.line)
		'Response.Write("<BR>LoadXML: " &  LoadXML)
		Response.Write("<BR>Reasonxx: " &  htPageDomCheck.parseError.reason)
		Response.End()
	end if

  	htPageDomCheck.selectSingleNode("SystemParameter/GIPconfig/" &funcName).text = saveText
  	
  	rtnVal = htPageDomCheck.save(LoadXMLCheck)
  	saveGIPApConfigText = rtnVal
end function
%>
