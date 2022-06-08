<%
	apCode = session("apCode")
	verDate = session("verDate")
	xid = request.queryString("id")
	xmlFile = request.queryString("xmlFile")
	
	response.write apCode & "<HR>"
	response.write verDate & "<HR>"
	response.write xid & "<HR>"
	response.write xmlFile & "<HR>"
'	response.end

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

	set vo = oxml.selectSingleNode("UseCase/Version[Date='"	& verDate & "']")
  if nullText(vo) = "" then
  	response.write "version not exist"
	response.end
  end if
  
  	eventDesc = nullText(vo.selectSingleNode("//TypicalCourseOfEvents/Event[@id='" & xid & "']"))
%>
<HTML>
<BODY>

<FORM method="POST" action="addContract.asp">
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=xid value="<%=xid%>">
<CENTER>
<TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="80%">
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>
XML檔案名稱：</TD>
<TD class="whitetablebg"><input name="htx_xmlFile" size="20" maxLength="20">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">說明：</TD>
<TD class="whitetablebg"><input name="htx_eventDesc" size="60" maxLength="60" value="<%=eventDesc%>">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">說明：</TD>
<TD class="whitetablebg"><input name="htx_dbDesc" size="50" maxLength="50">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">資料庫IP：</TD>
<TD class="whitetablebg"><input name="htx_dbIP" size="20" maxLength="20">
</TD>
</TR>
</TABLE>
</CENTER></FORM>
</BODY>
</HTML>
<%
	response.end

	set fxml = server.createObject("microsoft.XMLDOM")
	fxml.async = false
	fxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/HTSDSystem/expanded0.xml")
	xv = fxml.load(LoadXML)
  if fxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  fxml.parseError.line)
    Response.Write("<BR>Reason: " &  fxml.parseError.reason)
    Response.End()
  end if

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

	set xFieldNode = fxml.selectSingleNode("ExpandedUseCase/Version").cloneNode(true)
	
	xFieldNode.selectSingleNode("Date").text = date()
	xFieldNode.selectSingleNode("Author").text = session("userID")
	
	oxml.selectSingleNode("UseCase").insertBefore xFieldNode, oxml.selectSingleNode("UseCase/Version")
'	response.write oxml.xml
	oxml.save Server.MapPath("/HTSDSystem/UseCase/" & apCode & ".xml")
	response.redirect "ucVersion.asp?apCode=" & apCode



	
%>
