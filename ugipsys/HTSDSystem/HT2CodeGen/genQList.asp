<%@ Language=VBScript %>
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
	formFunction = "query"
'	session("ODBCDSN")="driver={SQL Server};server=61.13.76.20;UID=hometown;PWD=2986648;DATABASE=db921"
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

	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true	
	
	LoadXML = server.MapPath(".") & "\formSpec\" & formID & ".xml"
	debugPrint LoadXML & "<HR>"
	xv = htPageDom.load(LoadXML)
	response.write xv & "<HR>"
  if htPageDom.parseError.reason <> "" then 
    Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    Response.End()
  end if

pgPrefix = nullText(htPageDom.selectSingleNode("//htPage/HTProgPrefix"))
pgPath = request("programpath")
if pgPath = "" then pgPath = server.mapPath("genedCode/")
If right(pgPath,1) <> "\" then pgPath = pgPath & "\"

	Dim xSearchListItem(20,2)
	xItemCount = 0

    Set fs = CreateObject("scripting.filesystemobject")
'--xxxFORM.inc--------------------------
	set htFormDom = htPageDom.selectSingleNode("//pageSpec/formUI")
	set refModel = htPageDom.selectSingleNode("//htPage/resultSet")
'	response.write refModel.selectSingleNode("tableName").text & "<HR>"

	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

'	for each xCode in htFormDom.selectNodes("scriptCode")
'		xfout.write replace(xCode.text,chr(10),chr(13)&chr(10))
'	next

'--xxxQUERY.asp--------------------------
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"Query.asp")

	xfout.writeline cl & " Response.Expires = 0"
	xfout.writeline "HTProgCap=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageHead")) & cq
	xfout.writeline "HTProgFunc=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction")) & cq
	xfout.writeline "HTProgCode=""" & nullText(htPageDom.selectSingleNode("//htPage/HTProgCode")) & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempQuery1.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeline regOrgStr

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempQuery2.asp"))
	dumpTempFile
    xfin.Close

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempForm1.inc"))
	dumpTempFile
    xfin.Close
    
    	calendarFlag=false
    	for each param in refModel.selectNodes("paramList/param")
    		if nullText(param.selectSingleNode("clickYN"))="popDate" then calendarFlag=true : exit for 
    	next
    	if calendarFlag then
    		xfout.writeline "<object data=""../inc/calendar.htm"" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>"
		xfout.writeline "<INPUT TYPE=hidden name=CalendarTarget>"
    	end if
	for each x in htFormDom.selectSingleNode("pxHTML").childNodes		
		recursiveTag x
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempForm2.inc"))
	dumpTempFile
    xfin.Close

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempQuery3.asp"))
	dumpTempFile    		
    xfin.Close
    
    	if calendarFlag then
    		Set xfin = fs.opentextfile(Server.MapPath("Template1/TempCalendarSub.asp"))
			dumpTempFile    		
    		xfin.Close
    	end if
    	    
    xfout.Close



'--xxxLIST.asp--------------------------
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"List.asp")

	xfout.writeline cl & " Response.Expires = 0"
	xfout.writeline "HTProgCap=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageHead")) & cq
	xfout.writeline "HTProgFunc=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction")) & cq
	xfout.writeline "HTProgCode=""" & nullText(htPageDom.selectSingleNode("//htPage/HTProgCode")) & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempList1.asp"))
	dumpTempFile
    xfin.Close

	xSelect = refModel.selectSingleNode("sql/selectList").text
	xFrom = refModel.selectSingleNode("sql/fromList").text
	xrCount = 0
	for each param in refModel.selectNodes("fieldList/field[valueType='refLookup']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeID='" & param.selectSingleNode("refLookup").text & "'"
        SET RSLK=conn.execute(SQL)  
		xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & param.selectSingleNode("fieldName").text
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("CodeValueFld") & " = " & param.selectSingleNode("refField").text
		if not isNull(RSLK("CodeSrcFld")) then _
	    	xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
		xFrom = xFrom & ")"
	next

	xfout.writeline ct & "fSql = ""SELECT " & xSelect _
		& " FROM " & xFrom _
		& " WHERE " & refModel.selectSingleNode("sql/whereList").text & cq

	for each param in refModel.selectNodes("paramList/param")
		tagName = param.selectSingleNode("name").text
		paramCode = nullText(param.selectSingleNode("paramCode"))
		paramType = nullText(param.selectSingleNode("paramType"))
		paramKind = nullText(param.selectSingleNode("paramKind"))
		if paramKind = "range" then
			xfout.writeLine ct & "if request.form(""htx_" & paramCode & "S"") <> """" then"
		else
			xfout.writeLine ct & "if request.form(""htx_" & paramCode & """) <> """" then"
		end if
		select case paramKind
		  case "range"
			xfout.writeLine ct&ct & "rangeS = request.form(""htx_" & paramCode & "S"")"
			xfout.writeLine ct&ct & "rangeE = request.form(""htx_" & paramCode & "E"")"
			xfout.writeLine ct&ct & "if rangeE = """" then	rangeE=rangeS"
			xfout.writeLine ct&ct & "whereCondition = replace(""" & param.selectSingleNode("whereCondition").text _
				& """, ""{0}"", rangeS)"
			xfout.writeLine ct&ct & "whereCondition = replace(whereCondition, ""{1}"", rangeE)"
			xfout.writeLine ct&ct & "fSql = fSql & "" AND "" & whereCondition"
		  case else
			xfout.writeLine ct&ct & "whereCondition = replace(""" & param.selectSingleNode("whereCondition").text _
				& """, ""{0}"", request.form(""htx_" & paramCode & """) )"
			xfout.writeLine ct&ct & "fSql = fSql & "" AND "" & whereCondition"
		end select
		xfout.writeLine ct & "end if"		  	
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempList2.asp"))
	dumpTempFile
    xfin.Close
    
	for each param in refModel.selectNodes("fieldList/field[valueType != 'noop']")
		xfout.writeline ct & "<td align=center class=lightbluetable>" & nullText(param.selectSingleNode("fieldLabel")) & "</td>"
	next
    '---------------------------------------
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempList3.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeLine cl & "pKey = """"" 
	for each param in refModel.selectNodes("fieldList/field[isPrimaryKey='Y']")
		xfout.writeLine "pKey = pKey & ""&" & param.selectSingleNode("fieldName").text _
			& "="" & RSreg(""" & param.selectSingleNode("fieldName").text & """)"
	next	
	xfout.writeLine "if pKey<>"""" then  pKey = mid(pKey,2)" & cr

	for each param in refModel.selectNodes("fieldList/field[valueType != 'noop']")
		xfout.writeline ct & "<TD class=whitetablebg align=center><font size=2>" 
		xUrl = nullText(param.selectSingleNode("url"))
		if  xUrl <> "" then _
			xfout.writeLine ct & "<A href=""" & nullText(param.selectSingleNode("url")) & "?" & cl & "=pKey" & cr & """>"
		select case param.selectSingleNode("valueType").text
		  case "choice"
		  	xfout.writeLine ct & "<INPUT TYPE=radio name=""pkmain"" onClick=""setpKey ('" & cl & "=pKey" & cr & "')"">"
		  case "refLookup"
		  	xfout.writeLine cl & "=RSreg(""" & param.selectSingleNode("fieldName").text & """)" & cr
		  case "direct"
		  	xfout.writeLine cl & "=RSreg(""" & param.selectSingleNode("fieldName").text & """)" & cr
		  case "calc"
		  	xfout.writeLine cl & "=" & param.selectSingleNode("calc").text & cr
		end select
		if xUrl <> "" then	xfout.writeLine "</A>"
		xfout.writeLine "</font></td>"
	next
    '---------------------------
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempList4.asp"))
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

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempList5.asp"))
	dumpTempFile
    xfin.Close

    xc = 0
	for each param in refModel.selectNodes("funcButtonList/funcButton")
		xc = xc + 1
		xfout.writeLine ct & ct & "case " & xc & ": " _
			& param.selectSingleNode("action").text & " """ _
			& param.selectSingleNode("url").text & "?"" & gpKey"
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempList6.asp"))
	dumpTempFile
    xfin.Close
    xfout.Close

%>
</BODY>
</HTML>
