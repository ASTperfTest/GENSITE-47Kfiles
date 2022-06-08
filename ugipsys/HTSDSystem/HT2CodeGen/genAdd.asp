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
	formFunction = "add"
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

    Set fs = Server.CreateObject("scripting.filesystemobject")
'--xxxFORM.inc--------------------------
    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempFUForm1.inc"))
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

'--xxxADD.asp--------------------------
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"Add.asp")

	xfout.writeline cl & " Response.Expires = 0"
	xfout.writeline "HTProgCap=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageHead")) & cq
	xfout.writeline "HTProgFunc=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction")) & cq
	xfout.writeline "HTUploadPath=""" & nullText(htPageDom.selectSingleNode("//htPage/HTUploadPath")) & cq
	xfout.writeline "HTProgCode=""" & nullText(htPageDom.selectSingleNode("//htPage/HTProgCode")) & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempFUAdd1.asp"))
	dumpTempFile
    xfin.Close
    
	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processInit x.text
    next
    
    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempAdd1a.asp"))
	dumpTempFile
    xfin.Close

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processValid x.text
    next
    
    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempAdd1b.asp"))
	dumpTempFile
    xfin.Close
    

	xfout.writeline "<" & "!--#include file=""" & pgPrefix & "Form.inc""--" & ">"

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempAdd2.asp"))
	dumpTempFile
    xfin.Close

	showAidLinkList    

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempAdd2a.asp"))
	dumpTempFile
    xfin.Close

	masterRef = nullText(refModel.selectSingleNode("fieldList[masterTable='Y']/tableName"))
	xfout.writeline ct & "sql = ""INSERT INTO " & masterRef & "("""
	xfout.writeline ct & "sqlValue = "") VALUES("""

	for each x in htFormDom.selectNodes("//pxHTML//refField")
	    processInsert x.text
    next

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempFUAdd3.asp"))
	dumpTempFile
    xfin.Close
    xNewIdentiyName = nullText(refModel.selectSingleNode("fieldList[masterTable='Y']/field[identity='Y']/fieldName"))
	if xNewIdentiyName <> "" then
		xfout.writeline ct & "sql = ""set nocount on;""&sql&""; select @@IDENTITY as NewID"""
		xfout.writeline ct & "set RSx = conn.Execute(SQL)"
		xfout.writeline ct & "xNewIdentity = RSx(0)" & vbCRLF
	else
		xfout.writeline ct & "conn.Execute SQL" & vbCRLF
	end if	
	' --- process CloneMaster 
	xtTable = nullText(refModel.selectSingleNode("fieldList[cloneMaster='Y']/tableName"))
	if xtTable <> "" then
		if xNewIdentiyName <> "" then
			sqlWhere = xNewIdentiyName & " = "" & xNewIdentity"
		else
			sqlWhere = ""
			for each param in refModel.selectNodes("fieldList[tableName='" & masterRef & "']/field[isPrimaryKey='Y']")
				if sqlWhere <> "" then		sqlWhere = sqlWhere & " & "" AND "
				sqlWhere = sqlWhere & param.selectSingleNode("fieldName").text _
					& "="" & pkStr(xUpForm(""" & param.selectSingleNode("fieldName").text & """),"""")"
			next
		end if
		xfout.writeline ct & "xsql = ""SELECT * FROM " & masterRef & " WHERE " & sqlWhere
		xfout.writeline ct & "set RSmaster = conn.Execute(xSQL)" & vbCRLF
		
		xfout.writeline ct & "sql = ""INSERT INTO " & xtTable & "("""
		xfout.writeline ct & "sqlValue = "") VALUES("""
		
		for each x in refModel.selectNodes("fieldList[tableName='" & xtTable & "']/fkLink/fkFieldList")
			xfout.writeLine ct&ct&"sql = sql & """ & _
				nullText(x.selectSingleNode("myField")) & """ & "","""
			xfout.writeLine ct&ct&"sqlValue = sqlValue & pkStr(RSmaster(""" _
				& nullText(x.selectSingleNode("refField")) & """),"","")"
		next

		xfout.writeline
		
		for each x in htFormDom.selectNodes("//pxHTML//refField")
	    	processInsertTable x.text, xtTable
    	next

		xfout.writeline ct & "sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & "")"""
		xfout.writeline ct & "conn.Execute SQL" & vbCRLF
	end if
	' --- end process CloneMaster 

	xfout.writeline "end sub '---- doUpdateDB() ----" & vbCRLF
	xfout.writeline "sub showDoneBox()"
	
	doneURI = nullText(htPageDom.selectSingleNode("//htPage/doneURI"))
	xfout.writeLine ct & "doneURI= """ & doneURI & cq
	if doneURI <> "" AND xNewIdentiyName<> "" then _
		xfout.writeLine ct & "doneURI= doneURI & ""?" & xNewIdentiyName & "="" & xNewIdentity"

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempFUAdd3x.asp"))
	dumpTempFile
    xfin.Close
    xfout.Close

%>
</BODY>
</HTML>
