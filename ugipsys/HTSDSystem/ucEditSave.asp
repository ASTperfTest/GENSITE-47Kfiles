﻿<%
	Response.Expires = 0
'	response.write request("xmlBody")

	xmlString = request("xmlBody")
	set nxml = server.createObject("microsoft.XMLDOM")
	nxml.loadxml (xmlString)
	
	set oxml = session("xmlObj")
	set xnode = session("xmlEdit")
	set xparent = xnode.parentNode
	xparent.replaceChild nxml.documentElement.cloneNode(true), xnode
	
	oxml.save session("xmlPath")
	
%>
<script language=vbs>
'	alert window.opener.document.url
	window.opener.navigate window.opener.document.url
	window.close
</script>
<%

	response.end

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
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<CENTER>
<FORM name="reg" METHOD="POST" target="ucEditSave.asp">
<TEXTAREA name="xmlBody" rows="20" cols="80"><%=vo.xml%></TEXTAREA>
<INPUT type="submit" value="SAVE">

</FORM>
</body>
</html>