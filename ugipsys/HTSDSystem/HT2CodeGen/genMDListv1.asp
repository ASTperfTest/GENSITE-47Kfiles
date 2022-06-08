<%@ Language=VBScript %>
<% Response.Expires = 0 %>
<!--#Include file = "htCodeGen.inc" -->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<STYLE>
     td.fldLable {text-align:right; font-size:smaller;}
</STYLE>
</HEAD>
<BODY>
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


	formID = request("formID")
	progPath = request("progPath")
	if progPath <> "" then
		if left(progPath,1) = "/" then progPath = mid(progPath,2)
		progPath = replace(progPath,"/","\")
		progPath = progPath & "\"
	end if

	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true	
	
	LoadXML = server.MapPath(".") & "\formSpec\" & progPath & formID & ".xml"
	debugPrint LoadXML & "<HR>"
	xv = htPageDom.load(LoadXML)
  if htPageDom.parseError.reason <> "" then 
    Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    Response.Write("<BR>Reason: " &  htPageDom.parseError.reason)
    Response.End()
  end if

pgPrefix = nullText(htPageDom.selectSingleNode("//htPage/HTProgPrefix"))
progPath = nullText(htPageDom.selectSingleNode("//htPage/HTProgPath"))
if progPath = "" then 
	pgPath = server.mapPath("genedCode/")
else
	pgPath = server.mapPath(progPath)
end if
If right(pgPath,1) <> "\" then pgPath = pgPath & "\"

    Set fs = CreateObject("scripting.filesystemobject")
'--xxxFORM.inc--------------------------
	set htFormDom = htPageDom.selectSingleNode("//pageSpec/formUI")
	set refModel = htPageDom.selectSingleNode("//htForm/formModel[@id='" & htFormDom.selectSingleNode("@ref").text & "']")
'	response.write refModel.selectSingleNode("tableName").text & "<HR>"

	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

'	for each xCode in htFormDom.selectNodes("scriptCode")
'		xfout.write replace(xCode.text,chr(10),chr(13)&chr(10))
'	next

'--xxxLIST.asp--------------------------
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"List.asp")

	xfout.writeline cl & " Response.Expires = 0"
	xfout.writeline "HTProgCap=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageHead")) & cq
	xfout.writeline "HTProgFunc=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction")) & cq
	xfout.writeline "HTProgCode=""" & nullText(htPageDom.selectSingleNode("//htPage/HTProgCode")) & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempMDList1.asp"))
	dumpTempFile
    xfin.Close

	sqlWhere = ""
	urlPara = ""
	chkPara = ""
	xcount = 0
	masterRef = htFormDom.selectSingleNode("masterRef").text
	
	for each param in refModel.selectNodes("fieldList[tableName='" & masterRef & "']/field[isPrimaryKey='Y']")
		if xcount <> 0 then		sqlWhere = sqlWhere & " & "" AND "
		if urlPara <> "" then	urlPara = urlPara & "&"
		urlPara = urlPara & param.selectSingleNode("fieldName").text & "=" & cl & "=RSreg(""" & param.selectSingleNode("fieldName").text & """)" & cr
		chkPara=cl & "=RSreg(""" & param.selectSingleNode("fieldName").text & """)" & cr
		sqlWhere = sqlWhere & param.selectSingleNode("fieldName").text _
			& "="" & pkStr(request.queryString(""" & param.selectSingleNode("fieldName").text & """),"""")"
		xcount = xcount + 1
	next
	
	xSelect = "htx.*"
	xFrom = refModel.selectSingleNode("fieldList[masterTable='Y']/tableName").text & " AS htx"

	for each xfk in refModel.selectNodes("fieldList[fkLink/refTable='htx']")
		xAlias = xfk.selectSingleNode("fkLink/asAlias").text
		xFrom = "(" & xFrom & " " & xfk.selectSingleNode("fkLink/joinType").text & " JOIN " _
			& xfk.selectSingleNode("tableName").text & " AS " & xAlias & " ON " 
		for each xfkField in xfk.selectNodes("fkLink/fkFieldList")
			xFrom = xFrom & xAlias & "." & xfkField.selectSingleNode("myField").text & " = " _
				& "htx." _
				& xfkField.selectSingleNode("refField").text
		next
		xFrom = xFrom & ")"
		for each param in xfk.selectNodes("field")
			xSelect = xSelect & ", " & xAlias & "." & param.selectSingleNode("fieldName").text
		next
	next

	for each param in refModel.selectNodes("fieldList[masterTable='Y' or cloneMaster='Y']/field[inputType='file']")
		refField = param.selectSingleNode("fieldName").text
		xAlias = "xref" & refField
		xSelect = xSelect & ", " & xAlias & ".oldFileName AS fxr_" & refField
		xFrom = "(" & xFrom & " LEFT JOIN imageFile AS " & xAlias & " ON " _
			& xAlias & ".newFileName = htx." & refField & ")"
	next

	xfout.writeline ct & "sqlCom = ""SELECT " & xSelect & " FROM " & xFrom & " WHERE " & sqlWhere
	xfout.writeLine ct & "Set RSmaster = Conn.execute(sqlcom)"

	xfout.writeLine ct & "pKey = """"" 
	for each param in refModel.selectNodes("fieldList[tableName='" & masterRef & "']/field[isPrimaryKey='Y']")
		xfout.writeLine ct & "pKey = pKey & ""&" & param.selectSingleNode("fieldName").text _
			& "="" & RSmaster(""" & param.selectSingleNode("fieldName").text & """)"
	next	
	xfout.writeLine ct & "if pKey<>"""" then  pKey = mid(pKey,2)"


    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempMDList2.asp"))
	dumpTempFile
    xfin.Close

	for each x in htPageDom.selectNodes("//pageSpec/aidLinkList/Anchor")
		ckRight = nullText(x.selectSingleNode("checkRight"))
		pKeyStr = ""
		if nullText(x.selectSingleNode("AnchorWPK"))="Y" then _
			pKeyStr = "?" & cl & "=pKey" & cr
		if ckRight <> "" then _
			xfout.writeLine ct&ct & cl & "if (HTProgRight and " & ckRight & ")=" & ckRight & " then" & cr
		if x.selectSingleNode("AnchorType").text="Back" then
			xfout.writeLine ct&ct&ct & "<A href=""Javascript:window.history.back();" _
				& """ title=""" & nullText(x.selectSingleNode("AnchorDesc")) & """>" _
				& x.selectSingleNode("AnchorLabel").text & "</A> "
		else
			xfout.writeLine ct&ct&ct & "<A href=""" & x.selectSingleNode("AnchorURI").text _
				& pKeyStr & """ title=""" & nullText(x.selectSingleNode("AnchorDesc")) & """>" _
				& x.selectSingleNode("AnchorLabel").text & "</A>"
		end if
		if ckRight <> "" then _
			xfout.writeLine ct&ct & cl & "end if" & cr
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempMDList2a.asp"))
	dumpTempFile
    xfin.Close

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempForm1.inc"))
	dumpTempFile
    xfin.Close

	for each x in htFormDom.selectSingleNode("pxHTML").childNodes
		recursiveTag x
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempMDList2b.asp"))
	dumpTempFile
    xfin.Close

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    EditProcessInit x.text
    next
	xfout.writeLine "</script>"
	
	xfout.writeLine cl

	set xDetail = htFormDom.selectSingleNode("detailRow")
	detailRef = xDetail.selectSingleNode("detailRef").text

	xSelect = "dhtx.*"
	xFrom = detailRef & " AS dhtx"
	xrCount = 0

	for each param in xDetail.selectNodes("colSpec//refField")
	  xstr = param.text
	  inPos = instr(xstr,"/")
	  refTable = left(xstr,inPos-1)
	  refField = mid(xstr,inPos+1)
	  if refTable <> detailRef then
		if nullText(refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/fkLink/asAlias")) = "" then
			xrCount = xrCount + 1
			xAlias = "dref" & xrCount
			refModel.selectSingleNode("fieldList[tableName='"&refTable&"']/fkLink/asAlias").text=xAlias

			set xfk = refModel.selectSingleNode("fieldList[tableName='"&refTable&"']")
			xFrom = "(" & xFrom & " " & xfk.selectSingleNode("fkLink/joinType").text & " JOIN " _
				& xfk.selectSingleNode("tableName").text & " AS " & xAlias & " ON " 
			x1st = 0
			for each xfkField in xfk.selectNodes("fkLink/fkFieldList")
				if x1st > 0 then 	xFrom = xFrom & " AND "
				xFrom = xFrom & xAlias & "." & xfkField.selectSingleNode("myField").text & " = " _
					& xfk.selectSingleNode("fkLink/refTable").text & "." _
					& xfkField.selectSingleNode("refField").text
				x1st = x1st + 1
			next
			xFrom = xFrom & ")"
		else
			xAlias = refModel.selectSingleNode("fieldList[tableName='"&refTable&"']/fkLink/asAlias").text
		end if

        xAFldName = xAlias & "_" & refField
		xSelect = xSelect & ", " & xAlias & "." & refField & " AS " & xAFldName
        ' --- 把 detailRow 裡的 refField 換掉
        for each xd in xDetail.selectNodes("//colSpec//refField[text()='" _
        	& param.text & "']")
        	xd.text = xAlias & "/" & xAFldName
        next
        ' -----------------------------------
	  end if
	next
	xrCount = 0

	for each param in refModel.selectNodes("fieldList[tableName='" & detailRef & "']/field[refLookup!='']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeID='" & param.selectSingleNode("refLookup").text & "'"
        SET RSLK=conn.execute(SQL)
        response.write param.selectSingleNode("fieldName").text & "==>" & RSLK("CodeDisplayFld") & "<HR>"
		xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & param.selectSingleNode("fieldName").text
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("CodeValueFld") & " = dhtx." & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("CodeSrcFld")) then _
	    	xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
		xFrom = xFrom & ")"
	next

	xfout.writeline ct & "fSql = ""SELECT " & xSelect & cq & " _"
	xfout.writeline ct&ct & "& "" FROM " & xFrom & cq & " _"
	xfout.writeLine ct&ct & "& "" WHERE 1=1""" & " _"
	for each x in refModel.selectSingleNode("fieldList[tableName='" & detailRef & "']").selectNodes("fkLink/fkFieldList")
		xfout.writeLine ct&ct & "& "" AND dhtx." & x.selectSingleNode("myField").text & "="" & " _
			& "pkStr(RSmaster(""" & x.selectSingleNode("refField").text & """),"""")" & " _"
	next
	xOrderBy = " ORDER BY "
	if nullText(xDetail.selectSingleNode("divField")) <> "" then
		xOrderBy = xOrderBy & xDetail.selectSingleNode("divField").text & ", "
	end if
	for each x in xDetail.selectNodes("orderBy")
		xOrderBy = xOrderBy & x.text & ", "
	next
	if xOrderBy = " ORDER BY " 		then	xOrderBy = ""
	xfout.writeLine ct&ct & "& """ & xOrderBy & cq
	xfout.writeLine ct & "set RSlist = conn.execute(fSql)"

	xfout.writeLine cr
	
	for each param in xDetail.selectNodes("anchor")
	  select case nullText(param.selectSingleNode("type"))
		case "button"
			xfout.writeLine ct & "<INPUT TYPE=button VALUE=""" & param.selectSingleNode("funcLabel").text & """ onClick=""" _
				& param.selectSingleNode("action").text & "('" & param.selectSingleNode("url").text _
				& "?" & cl & "=pKey" & cr & "')"">"
		case "pTypeSet"
			xfout.writeLine ct & "<INPUT TYPE=button VALUE=""" & param.selectSingleNode("funcLabel").text & """ onClick=""" _
				& param.selectSingleNode("action").text & "('" & param.selectSingleNode("url").text _
				& "&" & cl & "=pKey" & cr & "')"">"				
	  end select
	next

    xfout.writeLine "<CENTER>"
    xfout.writeLine " <TABLE width=""95%"" cellspacing=""1"" cellpadding=""0"" class=""bg"">"
    xfout.writeLine " <tr align=""left"">"

	if nullText(xDetail.selectSingleNode("divLabel")) <> "" then
		xfout.writeline ct & "<td class=eTableLable>" & xDetail.selectSingleNode("divLabel").text & "</td>"
	end if
	for each param in xDetail.selectNodes("colSpec")
		xfout.writeline ct & "<td class=eTableLable>" & nullText(param.selectSingleNode("colLabel")) & "</td>"
	next
	xfout.writeLine " </tr>"
	
	xfout.writeLine cl
	xfout.writeLine ct & "while not RSlist.eof"
	xfout.writeLine ct & ct & "dpKey = """"" 
	for each param in refModel.selectNodes("fieldList[tableName='" & detailRef & "']/field[isPrimaryKey='Y']")
		xfout.writeLine ct & ct & "dpKey = dpKey & ""&" & param.selectSingleNode("fieldName").text _
			& "="" & RSlist(""" & param.selectSingleNode("fieldName").text & """)"
	next	
	xfout.writeLine ct & ct & "if dpKey<>"""" then  dpKey = mid(dpKey,2)"
	xfout.writeLine cr
    '---------------------------------------
	
	if nullText(xDetail.selectSingleNode("divField")) <> "" then
		xfout.writeline ct & "<TD class=eTableContent><font size=2>" 
	  	xfout.writeLine cl & "=RSlist(""" & xDetail.selectSingleNode("divField").text & """)" & cr
		xfout.writeLine "</font></td>"
	end if

	for each param in xDetail.selectNodes("colSpec")
		xfout.writeline ct & "<TD class=eTableContent><font size=2>" 
		xUrl = nullText(param.selectSingleNode("url"))
		if  xUrl <> "" then _
			xfout.writeLine ct & "<A href=""" & nullText(param.selectSingleNode("url")) & "?" & cl & "=dpKey" & cr & """>"
		processContent param.selectSingleNode("content")
		if xUrl <> "" then	xfout.writeLine "</A>"
		xfout.writeLine "</font></td>"
	next


    '---------------------------
    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempMDList4.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeLine ct & "<tr>"
    xfout.writeLine ct & ct & "<td width=""100%"" colspan=""2"" align=""center"">"
    xc = 0
	for each param in refModel.selectNodes("funcButtonList/funcButton")
		xc = xc + 1
		xfout.writeLine ct & ct & "<INPUT TYPE=button VALUE=""" & param.selectSingleNode("funcLabel").text _
			& """ onClick=""butAction(" & xc & ")"">"
	next
	xfout.writeLine ct & "</td></tr>"

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempMDList5.asp"))
	dumpTempFile
    xfin.Close
    xc = 0
	for each param in refModel.selectNodes("funcButtonList/funcButton")
		xc = xc + 1
		xfout.writeLine ct & ct & "case " & xc & ": " _
			& param.selectSingleNode("action").text & " """ _
			& param.selectSingleNode("url").text & "?"" & gpKey"
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempList6.asp"))
	dumpTempFile
    xfin.Close
    xfout.Close

sub processContent(xDom)
dim x
	if xDom.nodeName = "refField" then
		xstr = xDom.text
		inPos = instr(xstr,"/")
		refTable = left(xstr,inPos-1)
		refField = mid(xstr,inPos+1)
		on error resume next
		set xfNode = refModel.selectSingleNode("//field[fieldName='"&refField&"']")
		xfDataType = ""
		xfDataType = nullText(xfNode.selectSingleNode("dataType"))
		if xfDataType = "datetime" then
	  		xfout.writeLine cl & "=d7Date(RSlist(""" & refField & """))" & cr
		else
	  		xfout.writeLine cl & "=RSlist(""" & refField & """)" & cr
	  	end if
		exit sub
	end if
	if xDom.nodeName = "#comment" then	exit sub
	if xDom.nodeName = "#text" then
		xfout.writeLine xDom.text
		exit sub
	end if
	for each x in xDom.childNodes
		processContent x
	next
end sub
%>
</BODY>
</HTML>
