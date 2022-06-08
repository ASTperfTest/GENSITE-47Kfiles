<%
	formID = request("formID")

	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true	
	
	LoadXML = server.MapPath(".") & "\formSpec\" & formID & ".xml"
	debugPrint LoadXML & "<HR>"
	xv = htPageDom.load(LoadXML)
  if htPageDom.parseError.reason <> "" then 
    Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    Response.Write("<BR>Reason: " &  htPageDom.parseError.reason)
    Response.End()
  end if

'	set rptXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM")
'	rptXmlDoc.async = false
'	rptXmlDoc.setProperty("ServerHTTPRequest") = true	
	
'	LoadXML = server.MapPath(".") & "\formSpec\lbbdsd.xml"
'	debugPrint LoadXML & "<HR>"
'	xv = rptXmlDoc.load(LoadXML)
	
'  if rptXmlDoc.parseError.reason <> "" then 
'    Response.Write("XML parseError on line " &  rptXmlDoc.parseError.line)
'    Response.Write("<BR>Reason: " &  rptXmlDoc.parseError.reason)
'    Response.End()
'  end if

'	debugprint xv & "<HR>"
'	debugprint rptXmlDoc.xml
'	response.end

'	set formDom = rptXmlDoc.selectSingleNode("DataSchemaDef/dsTable[tableName='" & formID & "']")
	
'	debugprint formDom.xml
'	response.end

'	rptName = rptXmlDoc.selectSingleNode("RptTOC/RptDef[RptID='" & RptID & "']/RptName").text

'	set paramXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM")
'	paramXmlDoc.async = false
'	paramXmlDoc.setProperty("ServerHTTPRequest") = true
'	LoadXML = server.MapPath(".") & "\rqParam\" & rptName & ".xml"
'	debugPrint LoadXML & "<HR>"
'	xv = paramXmlDoc.load(LoadXML)
'	debugprint xv & "<HR>"
'	debugprint paramXmlDoc.xml
'	response.end
%>

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="setstyle.css">
<title><%=nullText(htPageDom.selectSingleNode("//pageSpec/pageHead"))%></title>
</head>
<body leftmargin="0">
<table border="0" width="100%">    
  <tr>    
    <td align="left">&nbsp;
    <span style="background:#0000A0;color:#FFFFFF;font-weight:bold;">
<%=nullText(htPageDom.selectSingleNode("//pageSpec/pageHead"))%></span></td>    
    <td align="right">    
<%	for each linkAnchor in htPageDom.selectNodes("//pageSpec/aidLinkList/link") %>
      &nbsp;&nbsp;<span><a href="<%=nullText(linkAnchor.selectSingleNode("uri"))%>"><%=nullText(linkAnchor.selectSingleNode("anchor"))%></a></span>
<%	next %>
    </td>
  </tr>    
</table>    

<center>
    <table width="90%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td class="c12-3">�� <%=nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction"))%> -</td>
      </tr>
</table>
<!-- �{���}�l --> 
                         
  <table width="95%" border="0" cellspacing="0" cellpadding="0">
 <form name="reg" method="POST">        
<object data="../inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style="position:absolute;top:100;left:230;visibility:hidden"></object>
<INPUT TYPE=hidden name=CalendarTarget>
<INPUT TYPE=hidden name=task>
    <tr>
        <td class="c12-1">&nbsp;</td>                                                  
    </tr>
  </table>
<CENTER><TABLE border=0 class=bluetable cellspacing=1 cellpadding=2 width="80%">
<%
	set htFormDom = htPageDom.selectSingleNode("//pageSpec/formUI")
	set refModel = htPageDom.selectSingleNode("//htForm/formModel[@id='" & htFormDom.selectSingleNode("@ref").text & "']")
'	response.write refModel.selectSingleNode("tableName").text & "<HR>"

dim baseTable

	set pxHTML = htFormDom.selectSingleNode("pxHTML")
	for each x in pxHTML.childNodes
		recursiveTag x
	next

	for each xCode in htFormDom.selectNodes("scriptCode")
		response.write xCode.text
	next

%>
</table>
</html>                         
<%
sub recursiveTag(xDom)
dim x
	if xDom.nodeName = "refField" then
		processParam(xDom.text)
		exit sub
	end if
	if xDom.nodeName = "#comment" then	exit sub
	if xDom.nodeName = "#text" then
		response.write xDom.text
		exit sub
	end if
	response.write "<" & xDom.nodeName
	for xi = 0 to xDom.attributes.length-1
		response.write " " & xDom.attributes.item(xi).nodeName & "=""" _
			& xDom.attributes.item(xi).text & """"
	next
	response.write ">"
	for each x in xDom.childNodes
'		response.write x.nodeName & "<BR>"
		recursiveTag x
	next
	response.write "</" & xDom.nodeName & ">" & vbCRLF
end sub


sub processFieldList(xflObj)
	xb = nullText(xflObj.selectSingleNode("tableName"))
	if xb <> "" then baseTable = xb
	for each xf in xflObj.selectNodes("field")
		processField xf
	next
end sub

sub processField(xfObj)
	response.write "<TR><TD class=""lightbluetable"" align=""right"">" & nullText(xfObj.selectSingleNode("fieldLabel")) & "�G</TD>" & vbCRLF
	response.write "<TD class=""whitetablebg""><refField>" & baseTable & "/" & xfObj.selectSingleNode("fieldName").text & "</refField>" & "</TD>" & vbCRLF
	response.write "</TR>" & vbCRLF
end sub

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
'	if not isNull(xNode) then
'	  if isObject(xNode) then
'		nullText = 
'	  end if
'	else
'		nullText = "aaa"
'	end if
end function

function ncTabChar(n)
	if cTabchar = "" then
		ncTabChar = ""
	else
		ncTabChar = String(n,cTabchar)
	end if
end function

sub debugPrint(xstr)
	response.write xstr
end sub

sub processParam(xstr)
'	response.write xstr
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
'	response.write xrefTable & "==>" & xrefField & "<BR>"
'	response.write "fieldList[tableName='" & xrefTable & "']/field[fieldName='" & refField & "']"
'	exit sub
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
		paramCode = param.selectSingleNode("fieldName").text
		paramType = param.selectSingleNode("dataType").text
		paramSize= nullText(param.selectSingleNode("dataLen"))
		if paramSize = "" then paramSize = 10
		if paramSize > 50 then paramSize = 50
		select case nullText(param.selectSingleNode("inputType"))
		  case ""
%>
      	<input name="htx_<%=paramCode%>" size="<%=paramSize%>">
<%
		  case "calc"
%>
      	<input name="htx_<%=paramCode%>" size="<%=paramSize%>" readonly="true">
<%
		  case "readonly"
%>
      	<input name="htx_<%=paramCode%>" size="<%=paramSize%>" readonly="true" class="rdonly">
<%
		  case "range"
%>
      	<input name="htx_<%=paramCode%>S" size="<%=paramSize%>"> �� <input name="htx_<%=paramCode%>E" size="<%=paramSize%>">
<%
		  case "varchar"
%>
      	<input name="htx_<%=paramCode%>" size="<%=paramSize%>">
<%
		  case "textarea"
%>
      	<textarea name="htx_<%=paramCode%>" rows="<%=nullText(param.selectSingleNode("rowsize"))%>" cols="<%=nullText(param.selectSingleNode("colsize"))%>">
      	</textarea>
<%
		  case "file"
%>
      	<input type="file" name="htx_<%=paramCode%>">
<%
		  case "smalldatetime"
%>
      	<input name="htx_<%=paramCode%>" size="<%=paramSize%>">
<%
		  case "multiCheckbox"
		  	refCode = param.selectSingleNode("refLookup").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
'		  	response.write refCode
		  	set optionList = rptXmlDoc.selectNodes("//dsTable[tableName='CodeMain']/instance/row[codeMetaID='"&refCode&"']")
			pdxc = ""
		  	for each optItem in optionList 
		  		if dynamicCode <> "" then _
		  			pdxc = replace(dynamicCode, "'mCode'", "'" & optItem.selectSingleNode("mCode").text & "'")
%>
		  	<input type="checkbox" name="htx_<%=paramCode%>" value="<%=optItem.selectSingleNode("mCode").text%>" <%=pdxc%>>
<%		  	response.write optItem.selectSingleNode("mValue").text & "�@"
		  	next
		  case "radio"
		  	refCode = param.selectSingleNode("refLookup").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
'		  	response.write refCode
		  	set optionList = rptXmlDoc.selectNodes("//dsTable[tableName='CodeMain']/instance/row[codeMetaID='"&refCode&"']")
			pdxc = ""
		  	for each optItem in optionList 
		  		if dynamicCode <> "" then _
		  			pdxc = replace(dynamicCode, "'mCode'", "'" & optItem.selectSingleNode("mCode").text & "'")
%>
		  	<input type="radio" name="htx_<%=paramCode%>" value="<%=optItem.selectSingleNode("mCode").text%>" <%=pdxc%>>
<%		  	response.write optItem.selectSingleNode("mValue").text & "&nbsp;&nbsp;"
		  	next
		  case "selection"
		  	refCode = param.selectSingleNode("refLookup").text
'		  	response.write refCode
		  	set optionList = rptXmlDoc.selectNodes("//dsTable[tableName='CodeMain']/instance/row[codeMetaID='"&refCode&"']")
%>
      <Select name="htx_<%=paramCode%>" size=1>
      		<option value="">�п��</option>
<%		  	for each optItem in optionList %>
		  	<option value="<%=optItem.selectSingleNode("mCode").text%>">
		  	<%=optItem.selectSingleNode("mValue").text%></option>
<%		  	next
			response.write "</select>" & vbCRLF
%>
<%
		  case "refSelect"
%>
      <Select name="htx_<%=paramCode%>" size=1>
      			<option value="">�п��</option>
		<% 	SQL= param.selectSingleNode("paramRefSQL").text		
			SET RSS=conn.execute(SQL)      			    	      			
			  	while not RSS.EOF %>
      				<option value=<%=RSS(0)%>><%=RSS(1)%></option>				  	
<%			  		RSS.movenext
			  	wend %>
      	</select>
<%
		end select
end sub


sub xprocessParam
		paramCode = param.selectSingleNode("fieldName").text
		paramType = param.selectSingleNode("dataType").text
		paramSize=10
		if paramType="varchar" then	paramSize = param.selectSingleNode("dataLen").text
		if paramType="char" then	paramSize = param.selectSingleNode("dataLen").text
		select case param.selectSingleNode("dataType").text
		  case "range"
%>
      	<input name="tfx_<%=paramCode%>S" size="<%=paramSize%>"> �� <input name="tfx_<%=paramCode%>E" size="<%=paramSize%>">
<%
		  case "varchar"
%>
      	<input name="tfx_<%=paramCode%>" size="<%=paramSize%>">
<%
		  case "smalldatetime"
%>
      	<input name="tfx_<%=paramCode%>" size="<%=paramSize%>">
<%
		  case "refLookup"
		  	refCode = param.selectSingleNode("refLookup").text
'		  	response.write refCode
		  	set optionList = rptXmlDoc.selectNodes("//dsTable[tableName='CodeMain']/instance/row[codeMetaID='"&refCode&"']")
%>
      <Select name="sfx_<%=paramCode%>" size=1>
      		<option value="">�п��</option>
<%		  	for each optItem in optionList %>
		  	<option value="<%=optItem.selectSingleNode("mCode").text%>">
		  	<%=optItem.selectSingleNode("mValue").text%></option>
<%		  	next
			response.write "</select>" & vbCRLF
%>
<%
		  case "refSelect"
%>
      <Select name="sfx_<%=paramCode%>" size=1>
      			<option value="">�п��</option>
		<% 	SQL= param.selectSingleNode("paramRefSQL").text		
			SET RSS=conn.execute(SQL)      			    	      			
			  	while not RSS.EOF %>
      				<option value=<%=RSS(0)%>><%=RSS(1)%></option>				  	
<%			  		RSS.movenext
			  	wend %>
      	</select>
<%
		  case "radio"
		  	set opList = param.selectNodes("availOption")
		  	for each opItem in opList
		  		checkedTrue = ""
		  		if nullText(opItem.selectSingleNode("default")) <> "" then checkedTrue = " CHECKED"
%>
      <INPUT TYPE="radio" name="rfx_<%=paramCode%>" VALUE="<%=opItem.selectSingleNode("value").text%>"<%=checkedTrue%>><%=opItem.selectSingleNode("lable").text%>
<%			next
		end select
end sub
%>