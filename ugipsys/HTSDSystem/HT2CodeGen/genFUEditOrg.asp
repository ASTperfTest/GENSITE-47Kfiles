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
'	response.end
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
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"FormE.inc", True)

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
'--xxxEDIT.asp--------------------------
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"Edit.asp")

	xfout.writeline cl & " Response.Expires = 0"
	xfout.writeline "HTProgCap=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageHead")) & cq
	xfout.writeline "HTProgFunc=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction")) & cq
	xfout.writeline "HTUploadPath=""" & nullText(htPageDom.selectSingleNode("//htPage/HTUploadPath")) & cq
	xfout.writeline "HTProgCode=""" & nullText(htPageDom.selectSingleNode("//htPage/HTProgCode")) & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempFUEdit1.asp"))
	dumpTempFile
    xfin.Close

	sqlWhere = ""
	urlPara = ""
	chkPara = ""
	xcount = 0
	
	for each param in refModel.selectNodes("fieldList/field[isPrimaryKey='Y']")
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

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempEdit2.asp"))
	dumpTempFile
    xfin.Close

	xSelect = "htx.*"
	xFrom = refModel.selectSingleNode("fieldList/tableName").text & " AS htx"

	for each param in refModel.selectNodes("fieldList/field[inputType='file']")
		refField = param.selectSingleNode("fieldName").text
		xAlias = "xref" & refField
		xSelect = xSelect & ", " & xAlias & ".oldFileName AS fxr_" & refField
		xFrom = "(" & xFrom & " LEFT JOIN imageFile AS " & xAlias & " ON " _
			& xAlias & ".newFileName = htx." & refField & ")"
	next

	xfout.writeline ct & "sqlCom = ""SELECT " & xSelect & " FROM " & xFrom & " WHERE " & sqlWhere
	xfout.writeLine ct & "Set RSreg = Conn.execute(sqlcom)"
	
	xfout.writeLine ct & "pKey = """"" 
	for each param in refModel.selectNodes("fieldList/field[isPrimaryKey='Y']")
		xfout.writeLine ct & "pKey = pKey & ""&" & param.selectSingleNode("fieldName").text _
			& "="" & RSreg(""" & param.selectSingleNode("fieldName").text & """)"
	next	
	xfout.writeLine ct & "if pKey<>"""" then  pKey = mid(pKey,2)"
	
	

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempEdit3.asp"))
	dumpTempFile
    xfin.Close

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processInit x.text
    next

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempFUEdit4.asp"))
	dumpTempFile
    xfin.Close

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processValid x.text
    next

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempEdit4a.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeline "<" & "!--#include file=""" & pgPrefix & "FormE.inc""--" & ">"

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempEdit5.asp"))
	dumpTempFile
    xfin.Close
   
	for each x in htPageDom.selectNodes("//pageSpec/aidLinkList/link/anchor")
		xfout.writeLine ct&ct&ct & "<A href=""" & x.selectSingleNode("url").text _
				& "?" & cl & "=pKey" & cr & """>" & x.selectSingleNode("funcLabel").text & "</A>"
	next
    

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempEdit5a.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeline ct & "sql = ""UPDATE " & refModel.selectSingleNode("fieldList/tableName").text & " SET """

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processUpdate x.text
    next
	xfout.writeLine ct & "sql = left(sql,len(sql)-1) & "" WHERE " & sqlWhere 

    Set xfin = fs.opentextfile(Server.MapPath("Template1/TempEdit6.asp"))
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

	if nullText(param.selectSingleNode("inputType")) = "imgfile"	then
		processImgFile xstr
		exit sub
	elseif nullText(param.selectSingleNode("inputType")) = "file"	then
		processAttFile xstr
		exit sub
	end if

'	xfout.writeLine ct&"IF request(""htx_" & refField & """) <> """" Then"
	select case paramType
	  case "calc"
	  case "integer"
			xfout.writeLine ct&ct&"sql = sql & """ & refField & "="" & " & "drn(""htx_" & refField & """)"
	  case else
			xfout.writeLine ct&ct&"sql = sql & """ & refField & "="" & " & "pkStr(xUpForm(""htx_" & refField & """),"","")"
	end select
'	xfout.writeLine ct&"END IF"
end sub

sub processImgFile(xstr)
'	response.write "<HR>" & xstr & "<HR>"
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	inputType = param.selectSingleNode("inputType").text

	xfout.writeLine ct&"IF xUpForm(""htImgActCK_" & refField & """) <> """" Then"
	xfout.writeLine ct&"  actCK = xUpForm(""htImgActCK_" & refField & """)"
	xfout.writeLine ct&"  if actCK=""editLogo"" OR actCK=""addLogo"" then"
	xfout.writeLine ct&ct&"fname = """""
	xfout.writeLine ct&ct&"For each xatt in xup.Attachments"
	xfout.writeLine ct&ct&"  if xatt.Name = ""htImg_" & refField & """ then"
	xfout.writeLine ct&ct&ct&"ofname = xatt.FileName"
	xfout.writeLine ct&ct&ct&"fnExt = """""
	xfout.writeLine ct&ct&ct&"if instr(ofname, ""."")>0 then	fnext=mid(ofname, instr(ofname, "".""))"
	xfout.writeLine ct&ct&ct&"tstr = now()"
	xfout.writeLine ct&ct&ct&"nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext"
	xfout.writeLine ct&ct&ct&"sql = sql & """ & refField & "="" & " & "pkStr(nfname,"","")"
	xfout.writeLine ct&ct&ct&"IF xUpForm(""hoImg_" & refField & """) <> """" Then _"
	xfout.writeLine ct&ct&ct&ct&"xup.DeleteFile apath & xUpForm(""hoImg_" & refField & """)"
	xfout.writeLine ct&ct&ct&"xatt.SaveFile apath & nfname, false"
	xfout.writeLine ct&ct&"  end if"
	xfout.writeLine ct&ct&"Next"
	xfout.writeLine ct&"  elseif actCK=""delLogo"" then"
	xfout.writeLine ct&ct&"xup.DeleteFile apath & xUpForm(""hoImg_" & refField & """)"
	xfout.writeLine ct&ct&"sql = sql & """ & refField & "=null,"""
	xfout.writeLine ct&"  end if"
	xfout.writeLine ct&"END IF"

end sub

sub processAttFile(xstr)
'	response.write "<HR>" & xstr & "<HR>"
	inPos = instr(xstr,"/")
	refTable = left(xstr,inPos-1)
	refField = mid(xstr,inPos+1)
	set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")
	inputType = param.selectSingleNode("inputType").text

	xfout.writeLine ct&"IF xUpForm(""htFileActCK_" & refField & """) <> """" Then"
	xfout.writeLine ct&"  actCK = xUpForm(""htFileActCK_" & refField & """)"
	xfout.writeLine ct&"  if actCK=""editLogo"" OR actCK=""addLogo"" then"
	xfout.writeLine ct&ct&"fname = """""
	xfout.writeLine ct&ct&"For each xatt in xup.Attachments"
	xfout.writeLine ct&ct&"  if xatt.Name = ""htFile_" & refField & """ then"
	xfout.writeLine ct&ct&ct&"ofname = xatt.FileName"
	xfout.writeLine ct&ct&ct&"fnExt = """""
	xfout.writeLine ct&ct&ct&"if instr(ofname, ""."")>0 then	fnext=mid(ofname, instr(ofname, "".""))"
	xfout.writeLine ct&ct&ct&"tstr = now()"
	xfout.writeLine ct&ct&ct&"nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext"
	xfout.writeLine ct&ct&ct&"sql = sql & """ & refField & "="" & " & "pkStr(nfname,"","")"
	xfout.writeLine ct&ct&ct&"IF xUpForm(""hoFile_" & refField & """) <> """" Then"
	xfout.writeLine ct&ct&ct&ct&"if xup.IsFileExist( apath & xUpForm(""hoFile_" & refField & """)) then _"
	xfout.writeLine ct&ct&ct&ct&ct&"xup.DeleteFile apath & xUpForm(""hoFile_" & refField & """)"
	xfout.writeLine ct&ct&ct&ct&"xsql = ""DELETE imageFile WHERE newFileName="" & pkStr(xUpForm(""hoFile_" & refField & """),"""")"
	xfout.writeLine ct&ct&ct&ct&"conn.execute xsql"
	xfout.writeLine ct&ct&ct&"END IF"
	xfout.writeLine ct&ct&ct&"xatt.SaveFile apath & nfname, false"
	xfout.writeLine ct&ct&ct&"xsql = ""INSERT INTO imageFile(newFileName, oldFileName) VALUES("" & pkStr(nfname,"","") & pkStr(ofname,"")"")"
	xfout.writeLine ct&ct&ct&"conn.execute xsql"
	xfout.writeLine ct&ct&"  end if"
	xfout.writeLine ct&ct&"Next"
	xfout.writeLine ct&"  elseif actCK=""delLogo"" then"
	xfout.writeLine ct&ct&"if xup.IsFileExist( apath & xUpForm(""hoFile_" & refField & """)) then _"
	xfout.writeLine ct&ct&ct&"xup.DeleteFile apath & xUpForm(""hoFile_" & refField & """)"
	xfout.writeLine ct&ct&"xsql = ""DELETE imageFile WHERE newFileName="" & pkStr(xUpForm(""hoFile_" & refField & """),"""")"
	xfout.writeLine ct&ct&"conn.execute xsql"
	xfout.writeLine ct&ct&"sql = sql & """ & refField & "=null,"""
	xfout.writeLine ct&"  end if"
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
	if nullText(param.selectSingleNode("calc")) <> "" then _
  		xfout.writeLine ct & param.selectSingleNode("calc").text
	if nullText(param.selectSingleNode("inputType")) <> "calc" then
	  select case nullText(param.selectSingleNode("inputType"))
	  	case "refRadio"
			xfout.writeLine ct & "initRadio ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
	  	case "radio"
			xfout.writeLine ct & "initRadio ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
	  	case "refCheckbox"
			xfout.writeLine ct & "initCheckbox ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
	  	case "checkbox"
			xfout.writeLine ct & "initCheckbox ""htx_" & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
	  	case "calc"
	  		xfout.writeLine ct & param.selectSingleNode("calc").text
		case "readonlydefault"	  	
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
  			xfout.writeLine	ct&"reg.htx_" & refField & ".readonly=""true"""	
			xfout.writeLine	ct&"reg.htx_" & refField & ".classname=""rdonly""" 	  		
	  	case "file"
			xfout.writeLine ct & "initAttFile """ & refField & """, " & cq & cl & "=qqRS(""" & refField & """)" & cr _
				& """, " & cq & cl & "=qqRS(""fxr_" & refField & """)" & cr & cq 
'			lhs = ct&"reg.hoFile_" & refField & ".value= "
'			rhs = cq & cl & "=qqRS(""fxr_" & refField & """)" & cr & cq
'  			xfout.writeLine lhs & rhs
	  	case "imgfile"
			xfout.writeLine ct & "initImgFile """ & refField & """," & cq & cl & "=qqRS(""" & refField & """)" & cr & cq 
'			lhs = ct&"document.all(""imgSrc_" & refField & """).src= "
'			rhs = cq & cl & "=HTUploadPath" & cr & cl & "=qqRS(""" & refField & """)" & cr & cq
'  			xfout.writeLine lhs & rhs
	  	case else
			lhs = ct&"reg.htx_" & refField & ".value= "
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
				& "?" & cl & "=pKey" & cr & "')"">"
	end select
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
		  case "calc"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly=""true"">"
		  case "readonlydefault"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""">"		  	
		  case "readonly"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly=""true"" class=""rdonly"">"
		  case "password"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" type=""password"">"
		  case "hidden"
		  	writeCode "<input type=""hidden"" name=""htx_"&paramCode&""">"
		  case "varchar"
		  	writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""">"
		  case "textarea"
		  	writeCode "<textarea name=""htx_"&paramCode&""" rows="""&nullText(param.selectSingleNode("rowsize")) _
		  		&""" cols="""&nullText(param.selectSingleNode("colsize"))&""">"
		  	writeCode "</textarea>"
		  case "file"
'		  	writeCode "<input type=""text"" name=""hoFile_" & paramCode & """ readonly=""true"" class=""rdonly"">"
'		  	writeCode "<input type=""file"" name=""htFile_" & paramCode & """>"

			writeCode "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
			writeCode ct & "<tr><td colspan=""2"">"
		  	writeCode ct&ct&"<input type=""file"" name=""htFile_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""hoFile_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""htFileActCK_" & paramCode & """>"
			writeCode ct & "</td></tr>"
			writeCode ct & "<tr>"
			writeCode ct & "<td width=""37%""><span id=""logo_" & paramCode & """ class=""rdonly""></span>"
		  	writeCode ct&ct&"<div id=""noLogo_" & paramCode & """ style=""color:red"">無附檔</div></td>"
			writeCode ct & "<td valign=""bottom"">"
		  	writeCode ct&ct&"<div id=""LbtnHide0_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 4)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""addLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/addimg.gif"" onClick=""VBS: addXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide1_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""chgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/chimg.gif"" onClick=""VBS: chgXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
 		  	writeCode ct&ct&cl&" if (HTProgRight and 16)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""delLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/delimg.gif"" onClick=""VBS: delXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide2_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""orgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/resetimg.gif"" onClick=""VBS: orgXFile '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
			writeCode ct & "</td></tr></table>"
		  case "imgfile"
			writeCode "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
			writeCode ct & "<tr><td colspan=""2"">"
		  	writeCode ct&ct&"<input type=""file"" name=""htImg_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""hoImg_" & paramCode & """>"
		  	writeCode ct&ct&"<input type=""hidden"" name=""htImgActCK_" & paramCode & """>"
			writeCode ct & "</td></tr>"
			writeCode ct & "<tr>"
			writeCode ct & "<td width=""37%""><img id=""logo_" & paramCode & """ src="""">"
		  	writeCode ct&ct&"<div id=""noLogo_" & paramCode & """ style=""color:red"">無圖片</div></td>"
			writeCode ct & "<td valign=""bottom"">"
		  	writeCode ct&ct&"<div id=""LbtnHide0_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 4)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""addLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/addimg.gif"" onClick=""VBS: addLogo '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide1_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""chgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/chimg.gif"" onClick=""VBS: chgLogo '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
 		  	writeCode ct&ct&cl&" if (HTProgRight and 16)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""delLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/delimg.gif"" onClick=""VBS: delLogo '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
		  	writeCode ct&ct&"<div id=""LbtnHide2_" & paramCode & """>"
 		  	writeCode ct&ct&cl&" if (HTProgRight and 8)<>0 then " & cr
		  	writeCode ct&ct&ct&"<img id=""orgLogo_" & paramCode & """ class=""hand"" src=""../pagestyle/resetimg.gif"" onClick=""VBS: orgLogo '" & paramCode & "'"">"
 		  	writeCode ct&ct&cl&" End if " & cr
		  	writeCode ct&ct&"</div>"
			writeCode ct & "</td></tr></table>"

'		  	writeCode "<img id=""imgSrc_" & paramCode & """ src="""">"
'		  	writeCode "<input type=""file"" name=""htImg_" & paramCode & """>"
		  case "popDate"
		  	calendarFlag = true
	  		writeCode "<input name=""htx_"&paramCode&""" size="""&paramSize&""" readonly onclick=""VBS: popCalendar 'htx_"&paramCode&"'"">"
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
		end select
end sub

%>
</BODY>
</HTML>
