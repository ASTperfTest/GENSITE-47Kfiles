<%@ CodePage = 65001 %>
<% Response.Expires = 0 
	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
'Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
conn.ConnectionTimeout=0
conn.CursorLocation = 3
conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	Server.ScriptTimeOut = 2000
	xmlURI = request("xml")
	formID = request("xml")
	createSchema = request("createSchema")

	set rptXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM")
	rptXmlDoc.async = false
	rptXmlDoc.setProperty("ServerHTTPRequest") = true	
	
	LoadXML = server.MapPath("/HTSDSystem/xmlSpec") & "\" & xmlURI & ".xml"
	debugPrint LoadXML & "<HR>"
	xv = rptXmlDoc.load(LoadXML)
	
  if rptXmlDoc.parseError.reason <> "" then 
    Response.Write("XML parseError on line " &  rptXmlDoc.parseError.line)
    Response.Write("<BR>Reason: " &  rptXmlDoc.parseError.reason)
    Response.End()
  end if

'	debugprint xv & "<HR>"
'	debugprint rptXmlDoc.xml
'	response.end

	set formDom = rptXmlDoc.selectSingleNode("DataSchemaDef/dsTable[tableName='" & formID & "']")
	
'	debugprint formDom.xml
'	response.end
%>

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="../PageStyle/system1.css">
<title>DB</title>
</head>
<body leftmargin="0">
<!-- 程式開始 --> 
<%
cvbCRLF = "<BR>" & vbCRLF
cTab = ""
	cSQL = "CREATE TABLE " & formID & " (" 
	set paramList = formDom.selectNodes("fieldList/field")
	for each param in paramList
		paramCode = param.selectSingleNode("fieldName").text
		paramType = param.selectSingleNode("dataType").text
		paramSize=10
		if paramType="varchar" then	paramSize = param.selectSingleNode("dataLen").text
		if paramType="char" then	paramSize = param.selectSingleNode("dataLen").text
		select case param.selectSingleNode("dataType").text
		  case "varchar"
		  	cSQL = cSQL & cvbCRLF & ctab & paramCode & " nvarchar(" & paramSize & ") COLLATE Chinese_Taiwan_Stroke_CS_AS"
		  case "char"
		  	cSQL = cSQL & cvbCRLF & ctab & paramCode & " nchar(" & paramSize & ") COLLATE Chinese_Taiwan_Stroke_CS_AS"
		  case else
		  	cSQL = cSQL & cvbCRLF & ctab & paramCode & " " & paramType
		end select
		if nullText(param.selectSingleNode("defaultValue")) <> "" then
			cSQL = cSQL & " DEFAULT '" & nullText(param.selectSingleNode("defaultValue")) & "'"
		end if
		if nullText(param.selectSingleNode("identity")) = "Y" then
			cSQL = cSQL & " IDENTITY"
		end if
		if nullText(param.selectSingleNode("canNull")) = "N" then
			cSQL = cSQL & " not null,"
		else
			cSQL = cSQL & " null,"
		end if
	    caption = nullText(param.selectSingleNode("fieldLabel"))
	    cSQL = cSQL & "  -- " & caption
	next 	
	cSQL = left(cSQL, len(cSQL)-len(caption)-6) & "  -- " & caption
	cSQL = cSQL & cvbCRLF & ")"
	response.write cSQL & "<BR><BR><BR>"
	if createSchema = "Y" then 		conn.execute cSQL

	cSQL = "ALTER TABLE " & formID & " ADD CONSTRAINT [PK_" & formID & "] PRIMARY KEY  NONCLUSTERED ("
	set paramList = formDom.selectNodes("fieldList/field[isPrimaryKey='Y']")
	for each param in paramList
		paramCode = param.selectSingleNode("fieldName").text
		cSQL = cSQL & paramCode & cvbCRLF & ","
	next
	cSQL = left(cSQL, len(cSQL)-1) & ")"
	response.write cSQL & "<HR>"
	if createSchema = "Y" then 		conn.execute cSQL

	
	set paramList = formDom.selectNodes("instance/row")
	for each param in paramList
'		response.write formDom.selectSingleNode("fieldList/field[isPrimaryKey='Y']/fieldName").text & "<HR>"
	  if nullText(param.selectSingleNode(formDom.selectSingleNode("fieldList/field[isPrimaryKey='Y']/fieldName").text)) <> "" then
		ifSQL = "INSERT INTO " & formID & "("
		ivSQL = " VALUES("
		for each fldValue in param.childNodes
'			response.write fldValue.nodeName & "==>" & fldValue.text & "<BR>"
			ifSQL = ifSQL & fldValue.nodeName & ","
			ivSQL = ivSQL & "'" & fldValue.text & "',"
		next
		iSQL = left(ifSQL,len(ifSQL)-1) & ")" & left(ivSQL,len(ivSQL)-1) & ")"
		response.write iSQL & "<HR>"
		conn.execute iSQL
	  end if
	next
%>
</body>
</html>                         

<%
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
%>
