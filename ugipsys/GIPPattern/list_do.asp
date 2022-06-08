<%@ CodePage=65001 Language="VBScript"%>
<%
'----條列清單

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
	qStr = request.queryString
	for each xf in request.form
		if request(xf)<>"" then	qStr = qStr & "&" & xf & "=" & server.URLEncode(request.form(xf))
	next
	LoadXML="http://ngipsys.hyweb.com.tw/GipPattern/list_do_xml.asp?specID=CtUnit"
	xv = oxml.load(LoadXML)
'	xv = oxml.load(server.mappath("test.xml"))
	if oxml.parseError.reason <> "" then 
		Response.Write("htPageDom parseError on line " &  oxml.parseError.line)
		Response.Write("<BR>Reason: " &  oxml.parseError.reason)
		Response.End()
	end if

	oxsl.load(server.mappath("xslGip/list.do PatternList.xsl"))

	Response.ContentType = "text/HTML" 
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	outString = replace(outString,"&amp;","&")
	response.write replace(outString,"&amp;","&")		
%>