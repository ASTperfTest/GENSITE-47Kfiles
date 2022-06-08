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
'	debugPrint LoadXML & "<HR>"
	xv = htPageDom.load(LoadXML)
  if htPageDom.parseError.reason <> "" then 
    Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    Response.Write("<BR>Reason: " &  htPageDom.parseError.reason)
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
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempFUForm1.inc"))
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"Form.inc", True)

	dumpTempFile
    xfin.Close

	set htFormDom = htPageDom.selectSingleNode("//pageSpec/formUI")
	set refModel = htPageDom.selectSingleNode("//htForm/formModel[@id='" & htFormDom.selectSingleNode("@ref").text & "']")
'	response.write refModel.selectSingleNode("tableName").text & "<HR>"

	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

   	calendarFlag=false
	for each x in htFormDom.selectSingleNode("pxHTML").childNodes
		recursiveTag x
	next
   	if calendarFlag then
   		xfout.writeline "<object data=""../inc/calendar.htm"" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>"
		xfout.writeline "<INPUT TYPE=hidden name=CalendarTarget>"
   	end if

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempForm2.inc"))
	dumpTempFile
    xfin.Close


	for each xCode in htFormDom.selectNodes("scriptCode")
		xfout.write replace(xCode.text,chr(10),chr(13)&chr(10))
	next
	
    	if calendarFlag then
    		Set xfin = fs.opentextfile(Server.MapPath("Template1/TempCalendarSub.asp"))
			dumpTempFile    		
    		xfin.Close
    	end if
    		
    xfout.Close

'--xxxADD.asp--------------------------
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"Add.asp")

	xfout.writeline cl & " Response.Expires = 0"
	xfout.writeline "HTProgCap=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageHead")) & cq
	xfout.writeline "HTProgFunc=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction")) & cq
	xfout.writeline "HTUploadPath=""" & nullText(htPageDom.selectSingleNode("//htPage/HTUploadPath")) & cq
	xfout.writeline "HTProgCode=""" & nullText(htPageDom.selectSingleNode("//htPage/HTProgCode")) & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempFUAdd1.asp"))
	dumpTempFile
    xfin.Close
    
	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processInit x.text
    next
    
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempAdd1a.asp"))
	dumpTempFile
    xfin.Close

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processValid x.text
    next
    
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempAdd1b.asp"))
	dumpTempFile
    xfin.Close
    

	xfout.writeline "<" & "!--#include file=""" & pgPrefix & "Form.inc""--" & ">"

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempAdd2.asp"))
	dumpTempFile
    xfin.Close

    GTable = refModel.selectSingleNode("fieldList[tableType='generic']/tableName").text
    DTable = refModel.selectSingleNode("fieldList[tableType='detail']/tableName").text
    GpKey = refModel.selectSingleNode("fieldList[tableType='generic']/field[identity='Y']/fieldName").text
    
    response.write "GTAble=" & GTable & "<HR>"
    response.write "DTable=" & DTable & "<HR>"
    response.write "GpKey=" & GpKey & "<HR>"

	xfout.writeline ct & "Gsql = ""INSERT INTO " & GTable & "("""
	xfout.writeline ct & "GsqlValue = "") VALUES("""
	xfout.writeline ct & "Dsql = ""INSERT INTO " & DTable & "("""
	xfout.writeline ct & "DsqlValue = "") VALUES("""

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processInsert x.text
    next

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempGSFUAdd3.asp"))
	dumpTempFile
    xfin.Close


	xfout.writeline ""
	xfout.writeline ct & "Gsql = left(Gsql,len(Gsql)-1) & left(GsqlValue,len(GsqlValue)-1) & "")"""
	xfout.writeline ct & "conn.Execute(Gsql)"
	xfout.writeline ct & "set RSmax = conn.execute(""SELECT MAX(" & GpKey & ") FROM " & GTable & """)"
	xfout.writeline ct & "DpKey = RSmax(0)"
	xfout.writeline ct & "Dsql = Dsql & """ & GpKey & """ & DsqlValue & DpKey & "")"""
	xfout.writeline ct & "conn.Execute(Dsql)"
	xfout.writeline ""

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempGSAdd3.asp"))
	dumpTempFile
    xfin.Close

    	if calendarFlag then
    		Set xfin = fs.opentextfile(Server.MapPath("Template1/TempCalendarSub.asp"))
			dumpTempFile    		
    		xfin.Close
    	end if
    	    
    xfout.Close
   
    
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

sub processInsert(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	GD = "G"
	if refTable=DTable then	GD="D"
		
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	paramType = param.selectSingleNode("dataType").text
	if lcase(right(nullText(param.selectSingleNode("inputType")),4)) = "file"	then	exit sub

	xfout.writeLine ct&"IF xUpForm(""htx_" & refField & """) <> """" Then"
	xfout.writeLine ct&ct&GD&"sql = " & GD & "sql & """ & refField & """ & "","""
	xfout.writeLine ct&ct&GD&"sqlValue = " & GD & "sqlValue & pkStr(xUpForm(""htx_" & refField & """),"","")"
	xfout.writeLine ct&"END IF"

end sub

sub processValid(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	it = nullText(param.selectSingleNode("inputType"))
	l  = nullText(param.selectSingleNode("dataLen"))
	if nullText(param.selectSingleNode("canNull")) = "N" then
	  if it<>"refRadio" AND it<>"radio" then
		xfout.writeLine ct&"IF reg.htx_" & refField & ".value = Empty Then"
		xfout.writeLine ct&ct&"MsgBox replace(nMsg,""{0}""," _
			& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
		xfout.writeLine ct&ct&"reg.htx_" & refField & ".focus"
		xfout.writeLine ct&ct&"exit sub"
		xfout.writeLine ct&"END IF"
	  end if
	end if
	dt = nullText(param.selectSingleNode("dataType"))
	if it = "imgfile" then
			xfout.writeLine ct&"IF reg.htImg_" & refField & ".value <> """" Then" 
			xfout.writeLine ct&ct&"xIMGname = reg.htImg_" & refField & ".value"
			xfout.writeLine ct&ct&"xFileType = """""
			xfout.writeLine ct&ct&"if instr(xIMGname, ""."")>0 then	xFileType=lcase(mid(xIMGname, instr(xIMGname, ""."")+1))"
			xfout.writeLine ct&ct&"IF xFileType<>""gif"" AND xFileType<>""jpg"" AND xFileType<>""jpeg"" then"
			xfout.writeLine ct&ct&ct&"MsgBox replace(pMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
			xfout.writeLine ct&ct&ct&"reg.htImg_" & refField & ".focus"
			xfout.writeLine ct&ct&ct&"exit sub"
			xfout.writeLine ct&ct&"END IF"
			xfout.writeLine ct&"END IF"
	elseif lcase(right(dt,4)) = "char" then
		if l<>"" AND (it="" OR it="textarea") then
			xfout.writeLine ct&"IF blen(reg.htx_" & refField & ".value) > " & l & " Then"
			xfout.writeLine ct&ct&"MsgBox replace(replace(lMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """),""{1}"","""&l&"""), 64, ""Sorry!"""
			xfout.writeLine ct&ct&"reg.htx_" & refField & ".focus"
			xfout.writeLine ct&ct&"exit sub"
			xfout.writeLine ct&"END IF"
		end if
	elseif lcase(right(dt,8)) = "datetime" then
			xfout.writeLine ct&"IF (reg.htx_" & refField & ".value <> """") AND (NOT isDate(reg.htx_" & refField & ".value)) Then" 
			xfout.writeLine ct&ct&"MsgBox replace(dMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
			xfout.writeLine ct&ct&"reg.htx_" & refField & ".focus"
			xfout.writeLine ct&ct&"exit sub"
			xfout.writeLine ct&"END IF"
	elseif lcase(left(dt,3)) = "int" then
			xfout.writeLine ct&"IF (reg.htx_" & refField & ".value <> """") AND (NOT isNumeric(reg.htx_" & refField & ".value)) Then" 
			xfout.writeLine ct&ct&"MsgBox replace(iMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
			xfout.writeLine ct&ct&"reg.htx_" & refField & ".focus"
			xfout.writeLine ct&ct&"exit sub"
			xfout.writeLine ct&"END IF"
	end if
end sub


sub processInit(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	if nullText(param.selectSingleNode("clientDefault/type")) = "" then	exit sub
	select case nullText(param.selectSingleNode("inputType"))
	  case "imgFile"
	  case "file"
	  case "refRadio"
		lhs = ct & "initRadio ""htx_" & refField & ""","
	  case "radio"
		lhs = ct & "initRadio ""htx_" & refField & ""","
	  case "refCheckbox"
		lhs = ct & "initCheckbox ""htx_" & refField & ""","
	  case "checkbox"
		lhs = ct & "initCheckbox ""htx_" & refField & ""","
	  case else
		lhs = ct&"reg.htx_" & refField & ".value= "
	end select
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
'			writeCode nullText(param.selectSingleNode("inputType"))
		select case nullText(param.selectSingleNode("inputType"))
		  case ""
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""">"
		  case "calc"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly=""true"">"
		  case "readonly"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly=""true"" class=""rdonly"">"
		  case "hidden"
		  	writeCode "<input type=""hidden"" name=""htx_"&paramCode&""">"
		  case "varchar"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""">"
		  case "textarea"
		  	writeCode "<textarea name=""htx_"&paramCode&""" rows="""&nullText(param.selectSingleNode("rowsize")) _
		  		&""" cols="""&nullText(param.selectSingleNode("colsize"))&""">"
		  	writeCode "</textarea>"
		  case "file"
		  	writeCode "<input type=""file"" name=""htFile_" & paramCode & """>"
		  case "imgfile"
		  	writeCode "<input type=""file"" name=""htImg_" & paramCode & """>"
		  case "popDate"
		  	calendarFlag = true
	  		writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"'"">"
'		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""">"
		  case "refCheckbox"
		  	refCode = param.selectSingleNode("refLookup").text
		  	tblRow = param.selectSingleNode("tblRow").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
		  	dynamicCode = replace(dynamicCode,chr(34),chr(34)&chr(34))
			writeCode "<table width=""100%"">"
			writeCode ct&ct&ct&cl&" pdxc = """""
			writeCode ct&ct&ct&ct&" tblRowCount = 0" & cr
			writeCode EnumerateCodeList(refCode)
			if dynamicCode <> "" then _
				writeCode ct&ct&ct&cl&" pdxc = replace("""& dynamicCode & """,""'mCode'"",""'""&RSS(0)&""'"")" & cr
			writeCode ct&ct&ct&cl& " if (tblRowCount mod " & tblRow & ") = 0 then  response.write ""<tr>""" & cr 
			writeCode ct&"<td><input type=""checkbox"" name=""htx_"&paramCode&""" value="""&cl&"=RSS(0)"&cr & """ " & cl&"=pdxc"&cr & ">" _
	  			&"<font size=2>"&cl&"=RSS(1)"&cr&"</font>"
			writeCode ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
				ct&ct&ct&"tblRowCount = tblRowCount + 1" & vbCrLf & _
	    		ct&ct&"wend"& cr
	    	writeCode "</table>"
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

		  case "selection"
			writeCode "<Select name=""htx_"&paramCode&""" size=1>"
      		writeCode "<option value="""">請選擇</option>"
		  	set optionList = param.selectNodes("item")
		  	for each optItem in optionList 
				writeCode "<option value="""&optItem.selectSingleNode("mCode").text & """>" _
		  			& optItem.selectSingleNode("mValue").text & "</option>"
		  	next
			writeCode "</select>" & vbCRLF

		  case "refSelect"
		  	refCode = param.selectSingleNode("refLookup").text
		  	dynamicCode = ""
		  		for each dc in param.selectNodes("dynamicBehavior")
		  			dynamicCode = " " & dc.selectSingleNode("event").text & "=""" _
		  				& dc.selectSingleNode("pCall").text & """"
		  		next
'		  	response.write refCode
'		  	set optionList = rptXmlDoc.selectNodes("//dsTable[tableName='CodeMain']/instance/row[codeMetaID='"&refCode&"']")
			writeCode "<Select name=""htx_"&paramCode&""" size=1>"
      		writeCode "<option value="""">請選擇</option>"
			writeCode EnumerateCodeList(refCode)
			writeCode ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"
			writeCode ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
	    		ct&ct&ct&"wend"& cr
			writeCode "</select>" & vbCRLF
		  case "ptypeselect"
		  	writeCode "<Select name=""htx_"&paramCode&""" size=1>"
			xfout.writeline "<" & "!--#include file=""../xml/pType.inc""--" & ">"
			writeCode "</select>" & vbCRLF			
		end select
end sub

%>
</BODY>
</HTML>
