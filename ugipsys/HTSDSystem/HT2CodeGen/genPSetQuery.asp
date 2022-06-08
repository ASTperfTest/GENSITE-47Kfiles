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
	response.write xv & "<HR>"
  if htPageDom.parseError.reason <> "" then 
    Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
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
	set htFormDom = htPageDom.selectSingleNode("//pageSpec/formUI")
	set refModel = htPageDom.selectSingleNode("//htPage/htForm/formModel")
'	response.write refModel.selectSingleNode("tableName").text & "<HR>"

	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

'--xxxListParam.inc---------------------
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"ListParam.inc")
    response.write pgPath&pgprefix&"ListParam.inc<HR>"
	xfout.writeline cl 
	xfout.writeline "sub xpCondition"

	for each uiField in htFormDom.selectNodes("pxHTML//refField")
		xstr = uiField.text
		inPos = instr(xstr,"/")
		refTable = left(xstr,inPos-1)
		refField = mid(xstr,inPos+1)
		set param = refModel.selectSingleNode("fieldList[tableName='" & refTable & "']/field[fieldName='" & refField & "']")

		tagName = param.selectSingleNode("fieldName").text
		paramCode = nullText(param.selectSingleNode("fieldName"))
		paramType = nullText(param.selectSingleNode("inputType"))
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
	
	xfout.writeline "end sub"
	xfout.writeline cr
    xfout.Close

'--xxxQUERY.asp--------------------------
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"Query.asp")

	xfout.writeline cl & " Response.Expires = 0"
	xfout.writeline "HTProgCap=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageHead")) & cq
	xfout.writeline "HTProgFunc=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction")) & cq
	xfout.writeline "HTProgCode=""" & nullText(htPageDom.selectSingleNode("//htPage/HTProgCode")) & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempSetQuery1.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeline regOrgStr

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempQuery2.asp"))
	dumpTempFile
    xfin.Close

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempForm1.inc"))
	dumpTempFile
    xfin.Close
    
    	calendarFlag=false
    	for each param in refModel.selectNodes("fieldList/field")
    		if nullText(param.selectSingleNode("inputType"))="popDate" then calendarFlag=true : exit for 
    	next
    	if calendarFlag then
    		xfout.writeline "<object data=""/inc/calendar.htm"" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>"
		xfout.writeline "<INPUT TYPE=hidden name=CalendarTarget>"
    	end if
	for each x in htFormDom.selectSingleNode("pxHTML").childNodes		
		recursiveQParam x
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempForm2.inc"))
	dumpTempFile
    xfin.Close

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempQuery3.asp"))
	dumpTempFile    		
    xfin.Close
    
	for each x in htPageDom.selectNodes("//pageSpec/aidLinkList/Anchor")
		ckRight = nullText(x.selectSingleNode("checkRight"))
		if ckRight <> "" then _
			xfout.writeLine ct&ct & cl & "if (HTProgRight and " & ckRight & ")=" & ckRight & " then" & cr
		if nullText(x.selectSingleNode("AnchorType")) = "Back" then
		  xfout.writeLine ct&ct&ct & "<A href=""Javascript:window.history.back();"">" _
				& x.selectSingleNode("AnchorLabel").text & "</A> "		
		else
		  xfout.writeLine ct&ct&ct & "<A href=""" & x.selectSingleNode("AnchorURI").text _
			& """ title=""" & nullText(x.selectSingleNode("AnchorDesc")) & """>" _
			& x.selectSingleNode("AnchorLabel").text & "</A>"
		end if
		if ckRight <> "" then _
			xfout.writeLine ct&ct & cl & "end if" & cr
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempQuery3a.asp"))
	dumpTempFile    		
    xfin.Close
    
    	if calendarFlag then
    		Set xfin = fs.opentextfile(Server.MapPath("Template0/TempCalendarSub.asp"))
			dumpTempFile    		
    		xfin.Close
    	end if
    	    
    xfout.Close

%>
</BODY>
</HTML>
