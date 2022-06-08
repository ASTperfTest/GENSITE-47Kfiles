<%@ CodePage = 65001 %>
<% 
'------ Modify History List (begin) ------
' 2006/1/1	92004/Chirs
'	1. if dataType is 'varchar' and dataLen is null, then generate ntext datatype
'	2. if inputType is range, generate xxx and xxx_RVH
'
'------ Modify History List (begin) ------
%>
<% Response.Expires = 0 
HTProgCap="單元資料定義"
HTProgCode="GE1T01"
HTProgPrefix="BaseDsd" %>
<!--#include virtual = "/inc/dbutil.inc" -->
<%

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

	set rptXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM")
	rptXmlDoc.async = false
	rptXmlDoc.setProperty("ServerHTTPRequest") = true
'	response.write "/site/" & session("mySiteID") & "/GipDSD/" & xmlURI & ".xml"
'	response.end
	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/" & xmlURI & ".xml")	
'	LoadXML = server.MapPath("/GipDSD/xmlSpec") & "\" & xmlURI & ".xml"
	debugPrint LoadXML & "<HR>"
	xv = rptXmlDoc.load(LoadXML)
	
  	if rptXmlDoc.parseError.reason <> "" then 
    		Response.Write("XML parseError on line " &  rptXmlDoc.parseError.line)
    		Response.Write("<BR>Reason: " &  rptXmlDoc.parseError.reason)
    		Response.End()
  	end if
	sqlcom="Select * from BaseDsd where ibaseDsd=" & pkStr(request.queryString("ibaseDsd"),"")
	Set RSmaster = Conn.execute(sqlcom)	
	if isNull(RSmaster("sBaseTableName")) then
		xTable = "CuDTx"+cStr(RSmaster("ibaseDsd"))
	else
		xTable = RSmaster("sBaseTableName")
	end if	
	set formDom = rptXmlDoc.selectSingleNode("DataSchemaDef/dsTable[tableName='" & xTable & "']")
	
%>

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="../PageStyle/system1.css">
<title>DB</title>
</head>
<body leftmargin="0">
<!-- 程式開始 --> 
<%
cSQL = ""
cvbCRLF = "<BR>" & vbCRLF
cTab = ""
	xn=0
	cSQL = "CREATE TABLE " & xTable & " (gicuitem int not null," 
	set paramList = formDom.selectNodes("fieldList/field")
	for each param in paramList
		paramCode = param.selectSingleNode("fieldName").text
		paramType = param.selectSingleNode("dataType").text
		paramSize=10
		if paramType="varchar" then	paramSize = param.selectSingleNode("dataLen").text
		if paramType="char" then	paramSize = param.selectSingleNode("dataLen").text
		select case param.selectSingleNode("dataType").text
		  case "varchar","text","nvarchar","ntext"
		    if paramSize="" then
			cSQL = cSQL & cvbCRLF & ctab & paramCode & " ntext COLLATE Chinese_Taiwan_Stroke_CS_AS"
		    else
			cSQL = cSQL & cvbCRLF & ctab & paramCode & " nvarchar(" & paramSize & ") COLLATE Chinese_Taiwan_Stroke_CS_AS"
		    end if
		  case "char","nchar"
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
			cSQL = cSQL & " null,"
		else
			cSQL = cSQL & " null,"
		end if
	    caption = nullText(param.selectSingleNode("fieldLabel"))
	    cSQL = cSQL & "  -- " & caption
	    xn=xn+1
	  if nullText(param.selectSingleNode("inputType"))="range" then
		paramCode = paramCode & "_RVH"
		select case param.selectSingleNode("dataType").text
		  case "varchar","text","nvarchar","ntext"
		    if paramSize="" then
			cSQL = cSQL & cvbCRLF & ctab & paramCode & " ntext COLLATE Chinese_Taiwan_Stroke_CS_AS"
		    else
			cSQL = cSQL & cvbCRLF & ctab & paramCode & " nvarchar(" & paramSize & ") COLLATE Chinese_Taiwan_Stroke_CS_AS"
		    end if
		  case "char","nchar"
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
			cSQL = cSQL & " null,"
		else
			cSQL = cSQL & " null,"
		end if
	    caption = nullText(param.selectSingleNode("fieldLabel")) & "上限值"
	    cSQL = cSQL & "  -- " & caption
	  end if
	next 	
	if xn>0 then 
		cSQL = left(cSQL, len(cSQL)-len(caption)-6) & "  -- " & caption
	else
		cSQL = Left(cSQL , len(cSQL)-1)
	end if
	cSQL = cSQL & cvbCRLF & ")"
	response.write cSQL & "<BR><BR><BR>"

	cSQL = "ALTER TABLE " & xTable & " ADD CONSTRAINT [PK_" & xTable & "] PRIMARY KEY  NONCLUSTERED (giCuItem)"
	response.write cSQL & "<HR>"
%>
</body>
</html>                         

<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

sub debugPrint(xstr)
	response.write xstr
end sub
%>
