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

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDList1.asp"))
	dumpTempFile
    xfin.Close

	sqlWhere = ""
	urlPara = ""
	chkPara = ""
	xcount = 0

	masterRef = htFormDom.selectSingleNode("masterRef").text
	set xDetail = htFormDom.selectSingleNode("detailRow")
	detailRef = xDetail.selectSingleNode("detailRef").text
	xfout.writeLine "masterRef=""" & masterRef & """"

	
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDEditList1.asp"))
	dumpTempFile
    xfin.Close
	

	for each param in refModel.selectNodes("fieldList[tableName='" & masterRef & "']/field[isPrimaryKey='Y']")
		if xcount <> 0 then		sqlWhere = sqlWhere & " & "" AND "
		if urlPara <> "" then	urlPara = urlPara & "&"
		urlPara = urlPara & param.selectSingleNode("fieldName").text & "=" & cl & "=RSreg(""" & param.selectSingleNode("fieldName").text & """)" & cr
		chkPara=cl & "=RSreg(""" & param.selectSingleNode("fieldName").text & """)" & cr
		sqlWhere = sqlWhere & param.selectSingleNode("fieldName").text _
			& "="" & pkStr(request.queryString(""" & param.selectSingleNode("fieldName").text & """),"""")"
		xcount = xcount + 1
	next
	sqlStr = ct & "SQL = ""DELETE FROM " & refModel.selectSingleNode("fieldList/tableName").text & " WHERE " & sqlWhere
	xfout.writeline sqlStr
	
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDEditList1a.asp"))
	dumpTempFile
    xfin.Close	
	
	sqlStr = ct & "sqlCom = ""SELECT * FROM " & refModel.selectSingleNode("fieldList/tableName").text & " WHERE " & sqlWhere
	xfout.writeline sqlStr
	xfout.writeLine ct & "Set RSmaster = Conn.execute(sqlcom)"

	xfout.writeLine ct & "mpKey = """"" 
	for each param in refModel.selectNodes("fieldList[tableName='" & masterRef & "']/field[isPrimaryKey='Y']")
		xfout.writeLine ct & "mpKey = mpKey & ""&" & param.selectSingleNode("fieldName").text _
			& "="" & RSmaster(""" & param.selectSingleNode("fieldName").text & """)"
	next	
	xfout.writeLine ct & "if mpKey<>"""" then  mpKey = mid(mpKey,2)"
	
	xfout.writeLine ct & "deleteFlag=false"
	xfout.writeLine ct & "SQLDeleteCheck=""Select * from " & detailRef & " WHERE " & sqlWhere
	xfout.writeLine ct & "SET RSDeleteCheck=conn.execute(SQLDeleteCheck)"
	xfout.writeLine ct & "if RSDeleteCheck.EOF then deleteFlag=true"		
	
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDEditList2a.asp"))
	dumpTempFile
    xfin.Close

    	calendarFlag=false
    	for each param in refModel.selectNodes("fieldList/field")  	
    		if nullText(param.selectSingleNode("inputType"))="popDate" then calendarFlag=true : exit for 
    	next
    	if calendarFlag then
    		xfout.writeline "<object data=""../inc/calendar.htm"" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>"
		xfout.writeline "<INPUT TYPE=hidden name=CalendarTarget>"
    	end if
    
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDEditList2b.asp"))
	dumpTempFile
    xfin.Close    

	for each x in htPageDom.selectNodes("//pageSpec/aidLinkList/anchor")
		xfout.writeLine ct&ct&ct & "<A href=""" & x.selectSingleNode("url").text _
				& "?" & cl & "=mpKey" & cr & """>" & x.selectSingleNode("funcLabel").text & "</A>"
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDList2a.asp"))
	dumpTempFile
    xfin.Close

	for each x in htFormDom.selectSingleNode("pxHTML").childNodes
		recursiveTag x
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDEditList2c.asp"))
	dumpTempFile
    xfin.Close 
	
	for each xCode in htFormDom.selectNodes("scriptCode")
		xfout.write replace(xCode.text,chr(10),chr(13)&chr(10))
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDList2c.asp"))
	dumpTempFile
    xfin.Close   

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDEditList2d.asp"))
	dumpTempFile
    xfin.Close

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processInit x.text
    next
 
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDEditList2e.asp"))
	dumpTempFile
    xfin.Close
    
	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processValid x.text
    next

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDEditList2f.asp"))
	dumpTempFile
    xfin.Close    
    
    
'---------
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDEditList3a.asp"))
	dumpTempFile
    xfin.Close  
    
	xfout.writeline ct & "sql = ""UPDATE " & refModel.selectSingleNode("fieldList/tableName").text & " SET """

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processUpdate x.text
    next
    
	xfout.writeLine ct & "sql = left(sql,len(sql)-1) & "" WHERE " & sqlWhere 
	

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDEditList6.asp"))
	dumpTempFile
    xfin.Close	
	    
'---------    
    	
	xfout.writeLine cl

	set xDetail = htFormDom.selectSingleNode("detailRow")
	detailRef = xDetail.selectSingleNode("detailRef").text

	xSelect = "htx.*"
	xFrom = xDetail.selectSingleNode("detailRef").text & " AS htx"
	xrCount = 0

	for each param in refModel.selectNodes("fieldList[tableName='" & detailRef & "']/field[valueType='refLookup']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeID='" & param.selectSingleNode("refLookup").text & "'"
        SET RSLK=conn.execute(SQL)
        response.write param.selectSingleNode("fieldName").text & "==>" & RSLK("CodeDisplayFld") & "<HR>"
		xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & param.selectSingleNode("fieldName").text
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("CodeValueFld") & " = htx." & param.selectSingleNode("refField").text
		if not isNull(RSLK("CodeSrcFld")) then _
	    	xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
		xFrom = xFrom & ")"
	next

	xfout.writeline ct & "fSql = ""SELECT " & xSelect _
		& " FROM " & xFrom & cq
	xfout.writeLine ct & "fSql = fSql & "" WHERE 1=1"""
	for each x in xDetail.selectNodes("fkLink")
		xfout.writeLine ct & "fSql = fSql & "" AND htx." & x.selectSingleNode("detailField").text & "="" & " _
			& "pkStr(RSmaster(""" & x.selectSingleNode("masterField").text & """),"""")"
	next
	xOrderBy = " ORDER BY "
	if nullText(xDetail.selectSingleNode("divField")) <> "" then
		xOrderBy = xOrderBy & xDetail.selectSingleNode("divField").text & ", "
	end if
	for each x in xDetail.selectNodes("orderBy")
		xOrderBy = xOrderBy & x.text & ", "
	next
	xfout.writeLine ct & "fSql = fSql & """ & xOrderBy & "1"""	
	xfout.writeLine ct & "set RSlist = conn.execute(fSql)"

	xfout.writeLine cr
	
	for each param in xDetail.selectNodes("anchor")
	  select case nullText(param.selectSingleNode("type"))
		case "button"
			xfout.writeLine ct & "<INPUT TYPE=button VALUE=""" & param.selectSingleNode("funcLabel").text & """ onClick=""" _
				& param.selectSingleNode("action").text & "('" & param.selectSingleNode("url").text _
				& "?" & cl & "=mpKey" & cr & "')"">"
	  end select
	next

    xfout.writeLine "<CENTER>"
    xfout.writeLine " <TABLE width=""95%"" cellspacing=""1"" cellpadding=""0"" class=""bg"">"
    xfout.writeLine " <tr align=""left"">"

	if nullText(xDetail.selectSingleNode("divLabel")) <> "" then
		xfout.writeline ct & "<td class=lightbluetable>" & xDetail.selectSingleNode("divLabel").text & "</td>"
	end if
	for each param in xDetail.selectNodes("colSpec")
		xfout.writeline ct & "<td class=lightbluetable>" & nullText(param.selectSingleNode("colLabel")) & "</td>"
	next
	xfout.writeLine " </tr>"
	
	xfout.writeLine cl
	xfout.writeLine ct & "while not RSlist.eof"
	xfout.writeLine ct & ct & "dpKey = """"" 
'response.write detailRef & "<br>"	
	for each param in refModel.selectNodes("fieldList[tableName='" & detailRef & "']/field[isPrimaryKey='Y']")
		xfout.writeLine ct & ct & "dpKey = dpKey & ""&" & param.selectSingleNode("fieldName").text _
			& "="" & RSlist(""" & param.selectSingleNode("fieldName").text & """)"
	next	
'response.write "xx=" & dpKey
'response.end	
	xfout.writeLine ct & ct & "if dpKey<>"""" then  dpKey = mid(dpKey,2)"
	xfout.writeLine cr
    '---------------------------------------
	
	if nullText(xDetail.selectSingleNode("divField")) <> "" then
		xfout.writeline ct & "<TD class=whitetablebg><font size=2>" 
	  	xfout.writeLine cl & "=RSlist(""" & xDetail.selectSingleNode("divField").text & """)" & cr
		xfout.writeLine "</font></td>"
	end if

	for each param in xDetail.selectNodes("colSpec")
		xfout.writeline ct & "<TD class=whitetablebg><font size=2>" 
		xUrl = nullText(param.selectSingleNode("url"))
		if  xUrl <> "" then _
			xfout.writeLine ct & "<A href=""" & nullText(param.selectSingleNode("url")) & "?" & cl & "=dpKey" & cr & """>"
		processContent param.selectSingleNode("content")
		if xUrl <> "" then	xfout.writeLine "</A>"
		xfout.writeLine "</font></td>"
	next


    '---------------------------
    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDList4.asp"))
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

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempMDEditList5.asp"))
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
    
    	if calendarFlag then
    		Set xfin = fs.opentextfile(Server.MapPath("Template1/TempCalendarSub.asp"))
			dumpTempFile    		
    		xfin.Close
    	end if
    	    
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

sub processContent(xDom)
dim x
	if xDom.nodeName = "refField" then
	  	xfout.writeLine cl & "=RSlist(""" & xDom.text & """)" & cr
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



sub recursiveTag(xDom)
dim x
	if xDom.nodeName = "refField" then
		processParam(xDom.text)
		exit sub
	end if
	if xDom.nodeName = "refAnchor" then
		processAnchor(xDom.text)
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

sub processUpdate(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	if nullText(param.selectSingleNode("isPrimaryKey")) = "Y" then	exit sub
	if nullText(param.selectSingleNode("inputType")) = "calc" then	exit sub
	paramType = param.selectSingleNode("dataType").text
'	xfout.writeLine ct&"IF request(""htx_" & refField & """) <> """" Then"
	select case paramType
	  case "calc"
	  case "integer"
			xfout.writeLine ct&ct&"sql = sql & """ & refField & "="" & " & "drn(""htx_" & refField & """)"
	  case else
			xfout.writeLine ct&ct&"sql = sql & """ & refField & "="" & " & "pkStr(request(""htx_" & refField & """),"","")"
	end select
'	xfout.writeLine ct&"END IF"
end sub

sub processInit(xstr)
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)	
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	if nullText(param.selectSingleNode("calc")) <> "" then _
  		xfout.writeLine ct & param.selectSingleNode("calc").text
	if nullText(param.selectSingleNode("inputType")) <> "calc" then
	  select case nullText(param.selectSingleNode("inputType"))
	  	case "refRadio"
			xfout.writeLine ct & "initRadio ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
	  	case "radio"
			xfout.writeLine ct & "initRadio ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
	  	case "calc"
	  		xfout.writeLine ct & param.selectSingleNode("calc").text
	  	case else
			lhs = ct&"document.all(""htx_" & refField & """).value= "
			rhs = cq & cl & "=qqRS(""" & refField & """)" & cr & cq
  			xfout.writeLine lhs & rhs
  	  end select
  	end if
end sub


sub processAnchor(xstr)
'	response.write "anchorList/anchor[funcLabel='" & xstr & "']<HR>"
	set param = refModel.selectSingleNode("anchorList/anchor[funcLabel='" & xstr & "']")
	select case nullText(param.selectSingleNode("type"))
		case "button"
			xfout.writeLine ct & "<INPUT TYPE=button VALUE=""" & xstr & """ onClick=""" _
				& param.selectSingleNode("action").text & "('" & param.selectSingleNode("url").text _
				& "?" & cl & "=mpKey" & cr & "')"">"
	end select
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
	if lcase(right(dt,4)) = "char" then
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
	elseif lcase(dt) = "integer" then
			xfout.writeLine ct&"IF (reg.htx_" & refField & ".value <> """") AND (NOT isNumeric(reg.htx_" & refField & ".value)) Then" 
			xfout.writeLine ct&ct&"MsgBox replace(iMsg,""{0}""," _
				& cq & param.selectSingleNode("fieldLabel").text & """), 64, ""Sorry!"""
			xfout.writeLine ct&ct&"reg.htx_" & refField & ".focus"
			xfout.writeLine ct&ct&"exit sub"
			xfout.writeLine ct&"END IF"
	end if
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
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""">"
		  case "hidden"
		  	writeCode "<input name=""htx_"&paramCode&""" type=""hidden"" size="""&paramSize&""">"
		  case "calc"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly=""true"">"
		  case "readonly"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly=""true"" class=""rdonly"">"
		  case "varchar"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""">"
		  case "textarea"
		  	writeCode "<textarea name=""htx_"&paramCode&""" rows="""&nullText(param.selectSingleNode("rowsize")) _
		  		&""" cols="""&nullText(param.selectSingleNode("colsize"))&""">"
		  	writeCode "</textarea>"
		  case "file"
		  	writeCode "<input type=""file"" name=""htx_" & paramCode & """>"
		  case "popDate"
		  	calendarFlag = true
	  		writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"'"">"
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
		end select
end sub

%>
</BODY>
</HTML>
