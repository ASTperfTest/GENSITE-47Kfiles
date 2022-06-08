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
	formFunction = "edit"

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
'	debugPrint LoadXML & "<HR>"
'	response.end
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
    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempFUForm1.inc"))
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

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempForm2.inc"))
	dumpTempFile
    xfin.Close


	for each xCode in htFormDom.selectNodes("scriptCode")
		xfout.write replace(xCode.text,chr(10),chr(13)&chr(10))
	next
	
    	if calendarFlag then
    		Set xfin = fs.opentextfile(Server.MapPath("Template0/TempCalendarSub.asp"))
			dumpTempFile    		
    		xfin.Close
    	end if
    	
    xfout.Close
'--xxxEDIT.asp--------------------------
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"Edit.asp")

	HTUploadPath = nullText(htPageDom.selectSingleNode("//htPage/HTUploadPath"))
	if HTUploadPath="" then		HTUploadPath = "/Public"
	xfout.writeline cl & " Response.Expires = 0"
	xfout.writeline "HTProgCap=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageHead")) & cq
	xfout.writeline "HTProgFunc=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction")) & cq
	xfout.writeline "HTUploadPath=""" & HTUploadPath & cq
	xfout.writeline "HTProgCode=""" & nullText(htPageDom.selectSingleNode("//htPage/HTProgCode")) & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempFUEdit1.asp"))
	dumpTempFile
    xfin.Close

	sqlWhere = ""
	xsqlWhere = ""
	urlPara = ""
	chkPara = ""
	xcount = 0
	
	for each param in refModel.selectNodes("fieldList[masterTable='Y']/field[isPrimaryKey='Y']")
		if xcount <> 0 then		
			sqlWhere = sqlWhere & " & "" AND "
			xsqlWhere = xsqlWhere & " & "" AND "
		end if
		if urlPara <> "" then	urlPara = urlPara & "&"
		urlPara = urlPara & param.selectSingleNode("fieldName").text & "=" & cl & "=RSreg(""" & param.selectSingleNode("fieldName").text & """)" & cr
		chkPara=cl & "=RSreg(""" & param.selectSingleNode("fieldName").text & """)" & cr
		sqlWhere = sqlWhere & param.selectSingleNode("fieldName").text _
			& "="" & pkStr(request.queryString(""" & param.selectSingleNode("fieldName").text & """),"""")"
		xsqlWhere = xsqlWhere & "htx." & param.selectSingleNode("fieldName").text _
			& "="" & pkStr(request.queryString(""" & param.selectSingleNode("fieldName").text & """),"""")"
		xcount = xcount + 1
	next
	sqlStr = ct & "SQL = ""DELETE FROM " & refModel.selectSingleNode("fieldList[masterTable='Y']/tableName").text & " WHERE " & sqlWhere
	xfout.writeline sqlStr

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempEdit2.asp"))
	dumpTempFile
    xfin.Close

	xSelect = "htx.*"
	xFrom = refModel.selectSingleNode("fieldList[masterTable='Y']/tableName").text & " AS htx"
	
	for each xfk in refModel.selectNodes("fieldList[fkLink]")
		xAlias = xfk.selectSingleNode("fkLink/asAlias").text
		xFrom = "(" & xFrom & " " & xfk.selectSingleNode("fkLink/joinType").text & " JOIN " _
			& xfk.selectSingleNode("tableName").text & " AS " & xAlias & " ON " 
		for each xfkField in xfk.selectNodes("fkLink/fkFieldList")
			xFrom = xFrom & xAlias & "." & xfkField.selectSingleNode("myField").text & " = " _
				& xfk.selectSingleNode("fkLink/refTable").text & "." _
				& xfkField.selectSingleNode("refField").text
		next
		xFrom = xFrom & ")"
		for each param in xfk.selectNodes("field")
			xSelect = xSelect & ", " & xAlias & "." & param.selectSingleNode("fieldName").text
		next
	next

	for each param in refModel.selectNodes("fieldList/field[inputType='file']")
		refField = param.selectSingleNode("fieldName").text
		xAlias = "xref" & refField
		xSelect = xSelect & ", " & xAlias & ".oldFileName AS fxr_" & refField
		xFrom = "(" & xFrom & " LEFT JOIN imageFile AS " & xAlias & " ON " _
			& xAlias & ".newFileName = htx." & refField & ")"
	next

	xfout.writeline ct & "sqlCom = ""SELECT " & xSelect & " FROM " & xFrom & " WHERE " & xsqlWhere
	xfout.writeLine ct & "Set RSreg = Conn.execute(sqlcom)"
	
	xfout.writeLine ct & "pKey = """"" 
	for each param in refModel.selectNodes("fieldList[masterTable='Y']/field[isPrimaryKey='Y']")
		xfout.writeLine ct & "pKey = pKey & ""&" & param.selectSingleNode("fieldName").text _
			& "="" & RSreg(""" & param.selectSingleNode("fieldName").text & """)"
	next	
	xfout.writeLine ct & "if pKey<>"""" then  pKey = mid(pKey,2)"
	
	

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempEdit3.asp"))
	dumpTempFile
    xfin.Close

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processInit x.text
    next

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempFUEdit4.asp"))
	dumpTempFile
    xfin.Close

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processValid x.text
    next

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempEdit4a.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeline "<" & "!--#include file=""" & pgPrefix & "FormE.inc""--" & ">"

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempEdit5.asp"))
	dumpTempFile
    xfin.Close
   
	showAidLinkList

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempEdit5a.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeline ct & "sql = ""UPDATE " & refModel.selectSingleNode("fieldList[masterTable='Y']/tableName").text & " SET """

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processUpdate x.text
    next
	xfout.writeLine ct & "sql = left(sql,len(sql)-1) & "" WHERE " & sqlWhere 

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempEdit6.asp"))
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
