<%@ Language=VBScript %>
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

	Dim xSearchListItem(20,2)
	xItemCount = 0

    Set fs = CreateObject("scripting.filesystemobject")
'--xxxFORM.inc--------------------------
	set htFormDom = htPageDom.selectSingleNode("//pageSpec/formUI")
	set refModel = htPageDom.selectSingleNode("//htPage/resultSet")
	set setModel = htPageDom.selectSingleNode("//htPage/setTable")
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

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempSetList1.asp"))
	dumpTempFile
    xfin.Close
    
	setTable = setModel.selectSingleNode("fieldList/tableName").text

	xfout.writeline ct&ct&ct&"sql = ""INSERT INTO " & setTable & "("" & session(""setPKeyField"")"
	xfout.writeline ct&ct&ct&"sqlValue = "") VALUES("" & session(""setPKeyValue"")"
	for each param in setModel.selectNodes("fieldList/field[setKey='Y']")
		fldName = param.selectSingleNode("fieldName").text
		xfout.writeline ct&ct&ct&"sql = sql & """ & "," & fldName & cq
		xfout.writeline ct&ct&ct&"sqlValue = sqlValue & "","" & pkStr(pickArray(xi),"""")"
	next
	for each param in setModel.selectNodes("fieldList/field[setNew='Y']")
		fldName = param.selectSingleNode("fieldName").text	
		if nullText(param.selectSingleNode("setDefault")) = "Y" then	
			xfout.writeline ct&ct&ct&ct&"sql = sql & """ & "," & fldName & cq
			xfout.writeline ct&ct&ct&ct&"sqlValue = sqlValue & "","" & pkStr(" & _
				param.selectSingleNode("setDefaultValue").text & ","""")"
		else	
			xfout.writeline ct&ct&ct&"if request(""" & fldName & """&pickArray(xi)) <> """" then"
			xfout.writeline ct&ct&ct&ct&"sql = sql & """ & "," & fldName & cq
			xfout.writeline ct&ct&ct&ct&"sqlValue = sqlValue & "","" & pkStr(request(""" & fldName & """&pickArray(xi)),"""")"
			xfout.writeline ct&ct&ct&"end if"
		end if
	next
    
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempSetList1a.asp"))
	dumpTempFile
    xfin.Close
    
    
    

	xSelect = refModel.selectSingleNode("sql/selectList").text
	xFrom = refModel.selectSingleNode("sql/fromList").text
	
	xFrom = "(" & xFrom & " LEFT JOIN " & setTable & " AS htxSet ON "" & session(""pkCondition"") & """
	
	fldName = setModel.selectSingleNode("fieldList/field[setKey='Y']/fieldName").text
	xSelect = xSelect & ",htxSet." & fldName & " AS htxSetExist"
	for each param in setModel.selectNodes("fieldList/field[setKey='Y']")
		fldName = param.selectSingleNode("fieldName").text
		xFrom = xFrom & " AND b." & fldName & "=htxSet." & fldName
	next
	xFrom = xFrom & ")"
	
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

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempSetList2.asp"))
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
	xfout.writeLine "pKeyValue = """"" 
	for each param in refModel.selectNodes("fieldList/field[isPrimaryKey='Y']")
		xfout.writeLine "pKey = pKey & ""&" & param.selectSingleNode("fieldName").text _
			& "="" & RSreg(""" & param.selectSingleNode("fieldName").text & """)"
		xfout.writeLine "pKeyValue = pKeyValue & ""&"" & RSreg(""" & param.selectSingleNode("fieldName").text & """)"
	next	
	xfout.writeLine "if pKey<>"""" then  pKey = mid(pKey,2)"
	xfout.writeLine "if pKeyValue<>"""" then  pKeyValue = mid(pKeyValue,2)" & cr

	for each param in refModel.selectNodes("fieldList/field[valueType != 'noop']")
		xfout.writeline ct & "<TD class=whitetablebg align=center><font size=2>" 
		xUrl = nullText(param.selectSingleNode("url"))
		if  xUrl <> "" then _
			xfout.writeLine ct & "<A href=""" & nullText(param.selectSingleNode("url")) & "?" & cl & "=pKey" & cr & """>"
		select case param.selectSingleNode("valueType").text
		  case "pickSet"
			xfout.writeLine ct & cl & "if isNull(RSreg(""htxSetExist"")) OR RSreg(""htxSetExist"") = """" then" & cr
		  	xfout.writeLine ct & ct & "<INPUT TYPE=checkbox name=""pickedItem"" value="""& cl & "=pKeyValue" & cr & """>"
			xfout.writeLine ct & cl & "end if" & cr
		  case "input"
			xfout.writeLine ct & cl & "if isNull(RSreg(""htxSetExist"")) OR RSreg(""htxSetExist"") = """" then" & cr
		  	xfout.writeLine ct & ct & "<INPUT TYPE=text name=""" _
		  		& param.selectSingleNode("fieldName").text&cl & "=pKeyValue" & cr & """ size=""" _
		  		& param.selectSingleNode("dataLen").text& """>"
			xfout.writeLine ct & cl & "end if" & cr
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
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempSetList4.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeLine ct & "<tr>"
    xfout.writeLine ct & ct & "<td width=""100%"" colspan=""2"" align=""center"">"

    xfout.writeLine ct & ct & "<input type=submit value =""確　定"" class=""cbutton"" name=""submitTask"">"
    xfout.writeLine ct & ct & "<input type=button value =""回前頁"" class=""cbutton"" onClick=""Vbs:history.back"">"

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

	response.end
    
    
sub dumpTempFile()
    Do While Not xfin.AtEndOfStream
        xinStr = xfin.readline
        xfout.writeline xinStr
    Loop
end sub

sub writeCode(xstr)
	response.write xstr
	xfout.writeline xstr
end sub

sub writePart(xstr)
	response.write xstr
	xfout.write xstr
end sub

Function EnumerateCodeList(codeID)
	SQL="Select * from CodeMetaDef where codeID='" & codeID & "'"
        SET RSLK=conn.execute(SQL)
	str=""
	if not RSLK.EOF then
	  if isNull(RSLK("CodeSortFld")) then
		if isNull(RSLK("CodeSrcFld")) then
	    	str=ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & """" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
		    	ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
		else	
	    	str=ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'""" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
			    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
		end if          
	  else
		if isNull(RSLK("CodeSrcFld")) then
	    	str=ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL Order by " & RSLK("CodeSortFld") & """" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
			    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
	    else	
	    	str=ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL AND " & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "' Order by " & RSLK("CodeSortFld") & """" & VBCRLF & _
		    	ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
			    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF
		end if
	  end if
	end if
	EnumerateCodeList=str
End Function

sub recursiveTag(xDom)
dim x
	if xDom.nodeName = "refField" then	
		processParam(xDom.text)
		exit sub
	end if
	if xDom.nodeName = "#comment" then	exit sub
	if xDom.nodeName = "#text" then
		writePart xDom.text
		exit sub
	end if
	writePart "<" & xDom.nodeName
	for xi = 0 to xDom.attributes.length-1
		writePart " " & xDom.attributes.item(xi).nodeName & "=""" _
			& xDom.attributes.item(xi).text & """"
	next
	writePart ">"
	for each x in xDom.childNodes
		recursiveTag x
	next
	writeCode "</" & xDom.nodeName & ">"
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

sub processInit(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	if nullText(param.selectSingleNode("clientDefault/type")) = "" then	exit sub
	lhs = ct&"reg.htx_" & refField & ".value= "
	select case param.selectSingleNode("clientDefault/type").text
	  case "value"
	  	rhs = cq & param.selectSingleNode("clientDefault/set").text & cq
	  case "clientFunc"
	  	rhs = param.selectSingleNode("clientDefault/set").text
	  case "serverFunc"
	  	rhs = cq & cl & "=" & param.selectSingleNode("clientDefault/set").text & cr & cq
	  case "session"
	  	rhs = cq & cl & "=session(" & param.selectSingleNode("clientDefault/set").text & ")" & cr & cq
	end select
  	xfout.writeLine lhs & rhs
end sub

sub processParam(xstr)
	refField = xstr
'response.write refField	
'	response.write xrefTable & "==>" & xrefField & "<BR>"
'	response.write "fieldList[tableName='" & xrefTable & "']/field[fieldName='" & refField & "']"
'	exit sub
	set param = refModel.selectSingleNode("paramList/param[name='" & refField & "']")	
		paramCode = param.selectSingleNode("paramCode").text
		paramType = param.selectSingleNode("paramType").text
		paramSize= nullText(param.selectSingleNode("paramSize"))
		if paramSize = "" then paramSize = 10
		if paramSize > 50 then paramSize = 50
		select case param.selectSingleNode("paramKind").text
		  case "range"
		  	if param.selectSingleNode("clickYN").text="popDate" then
		  		writeCode "<input name=""htx_"&paramCode&"S"" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"S'""> ～ "
		  		writeCode "<input name=""htx_"&paramCode&"E"" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"E'"">"
		  	else
		  		writeCode "<input name=""htx_"&paramCode&"S"" size="""&paramSize&"""> ～ "
		  		writeCode "<input name=""htx_"&paramCode&"E"" size="""&paramSize&""">"		  	
		  	end if
		  case "value"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""">"
		  case "refSelect"
		  	refCode = param.selectSingleNode("refLookup").text	  	
			writeCode "<Select name=""htx_"&paramCode&""" size=1>"			
      		writeCode "<option value="""">請選擇</option>"
			writeCode EnumerateCodeList(refCode)
			writeCode ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr
			writeCode "</select>" & vbCRLF		
		  case "selection"
			writeCode "<Select name=""htx_"&paramCode&""" size=1>"
      		writeCode "<option value="""">請選擇</option>"
		  	set optionList = param.selectNodes("item")
		  	for each optItem in optionList 
				writeCode "<option value="""&optItem.selectSingleNode("mCode").text & """>" _
		  			& optItem.selectSingleNode("mValue").text & "</option>"
		  	next
			writeCode "</select>" & vbCRLF
		  case "radio"
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	set optionList = param.selectNodes("item")
			pdxc = ""
		  	for each optItem in optionList 
		  		if dynamicCode <> "" then _
		  			pdxc = replace(dynamicCode, "'mCode'", "'" & optItem.selectSingleNode("mCode").text & "'")
				writeCode "<input type=""radio"" name=""htx_"&paramCode&""" value="""&optItem.selectSingleNode("mCode").text & """ " & pdxc & ">"
		  		writeCode optItem.selectSingleNode("mValue").text & "&nbsp;&nbsp;"
		  	next
		  case "refRadio"
		  	refCode = param.selectSingleNode("refLookup").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	dynamicCode = replace(dynamicCode,chr(34),chr(34)&chr(34))
			writeCode ct&ct&ct&cl&" pdxc = """"" & cr
			writeCode EnumerateCodeList(refCode)
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
			writeCode "<input type=""radio"" name=""htx_"&paramCode&""" value="""&cl&"=RSS(0)"&cr & """ " & cl&"=pdxc"&cr & ">"
	  		writeCode ct&ct&ct&cl&"=RSS(1)"&cr& "&nbsp;&nbsp;"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr
		  case "refCheckbox"
		  	refCode = param.selectSingleNode("refLookup").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	dynamicCode = replace(dynamicCode,chr(34),chr(34)&chr(34))
			writeCode ct&ct&ct&cl&" pdxc = """"" & cr
			writeCode EnumerateCodeList(refCode)
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
			writeCode "<input type=""checkbox"" name=""htx_"&paramCode&""" value="""&cl&"=RSS(0)"&cr & """ " & cl&"=pdxc"&cr & ">"
	  		writeCode ct&ct&ct&cl&"=RSS(1)"&cr& "　"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr
		end select
end sub

%>
</BODY>
</HTML>
