<% HTProgCap="AP"
HTProgCode="HT090"
HTProgPrefix="AP" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<%
	apCode = request.queryString("apCode")

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true
	LoadXML = server.mappath("/HTSDSystem/UseCase/00APlist.xml")
	xv = oxml.load(LoadXML)
  if oxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
%>
<HR>
<A href="genAPReport.asp">Gen AP Report</A>
<%    
    Response.End()
  end if

	set oxsl = server.createObject("microsoft.XMLDOM")
	  oxsl.async = false
	xv = oxsl.load(server.mappath("/HTSDSystem/APreport.xsl"))
'	response.write vo.transformNode(oxsl)
	Response.ContentType = "text/HTML" 
	response.write replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	response.end

%>
