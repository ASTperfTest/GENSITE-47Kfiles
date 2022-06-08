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
		xv = oxsl.load(server.mappath("/HTSDSystem/ucReport.xsl"))
'	response.write vo.transformNode(oxsl)
	Response.ContentType = "text/HTML" 
	response.write replace(vo.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	response.end



	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<CENTER>
<TABLE border=0 class=bluetable cellspacing=1 cellpadding=2 width="90%">
	<TR><TD class=lightbluetable align=right width="20%">類別/AP代碼</TD>
	  <TD class=whitetablebg>&nbsp;
	  <%=nullText(oxml.selectSingleNode("UseCase/APCat"))%> - 
	  <%=nullText(oxml.selectSingleNode("UseCase/Code"))%> 
		　　　　　　　　<A href="ucXML.asp?apCode=<%=apCode%>">View XML</A>
	  </TD></TR>
	<TR><TD class=lightbluetable align=right width="20%">模組名稱</TD>
	  <TD class=whitetablebg>&nbsp;
	  <%=nullText(oxml.selectSingleNode("UseCase/Name"))%>
	  </TD></TR>
	<TR><TD class=lightbluetable align=right width="20%">版本</TD>
	  <TD class=whitetablebg align=center>
<TABLE border=0 class=bluetable cellspacing=1 cellpadding=2 width="90%">
	<TR>
		<TD class=lightbluetable align=center>日期</TD>
		<TD class=lightbluetable align=center>類別</TD>
		<TD class=lightbluetable align=center>作者</TD>
	</tr>
<%	for each vo in oxml.selectNodes("UseCase/Version") 
		select case vo.selectSingleNode("Kind").text
		  case "Expanded"
		  	xhref = "ucExpanded.asp"
		  case "HighLevel"
		  	xhref = "ucHighLevel.asp"
		end select
		xhref = xhref & "?apCode=" & apCode & "&verDate="
%>
	<TR>
	  <TD class=whitetablebg align=center>
	  	<A href="<%=xhref%><%=nullText(vo.selectSingleNode("Date"))%>">
	  	<%=nullText(vo.selectSingleNode("Date"))%></A>
	  </TD>
	  <TD class=whitetablebg align=center>
	  	<%=nullText(vo.selectSingleNode("Kind"))%>
	  </TD>
	  <TD class=whitetablebg align=center>
	  	<%=nullText(vo.selectSingleNode("Author"))%>
	  </TD>
<%	next	%>
	</TABLE>
	  </TD></TR>
</TABLE>
</body>
</html>