<%
function checkGIPconfig(funcName){
		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysPara.xml"
		xv = htPageDom.load(LoadXML)
  		if htPageDom.parseError.reason <> "" then 
    		Response.Write("sysPara.xml parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  		end if
  	checkGIPconfig = false
  	if UCASE(nullText(htPageDom.selectSingleNode("SystemParameter/GIPconfig/" &funcName)))="Y" then	_
		checkGIPconfig = true
}

%>
