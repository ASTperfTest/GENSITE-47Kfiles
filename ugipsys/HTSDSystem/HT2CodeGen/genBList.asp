<%@ Language=VBScript %>
<!--#Include file = "htCodeBGen.inc" -->
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
	formFunction = "list"
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
	set htFormDom = htPageDom.selectSingleNode("//pageSpec")
	set refModel = htPageDom.selectSingleNode("//htPage/resultSet")
	set xDetail = htFormDom.selectSingleNode("detailRow")
'	response.write refModel.selectSingleNode("tableName").text & "<HR>"

	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

'--xxxListParam.inc---------------------
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"ListParam.inc")
    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempListParam0.inc"))
	dumpTempFile
    xfin.Close
    xfout.Close


'--xxxLIST.asp--------------------------
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"List.asp")

	xfout.writeline cl & " Response.Expires = 0"
	xfout.writeline "HTProgCap=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageHead")) & cq
	xfout.writeline "HTProgFunc=""" & nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction")) & cq
	xfout.writeline "HTProgCode=""" & nullText(htPageDom.selectSingleNode("//htPage/HTProgCode")) & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr
	xfout.writeline "<!--#INCLUDE FILE=""" & pgPrefix & "ListParam.inc"" -->"


    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempBList1.asp"))
	dumpTempFile
    xfin.Close

	masterTableName = refModel.selectSingleNode("fieldList[masterTable='Y']/tableName").text
	xfout.writeline ct&ct & "upSql = ""UPDATE " & masterTableName & " SET "" _"
	pad1st = ""
	uxSet = nullText(htPageDom.selectSingleNode("//htPage/updateSet"))
	if uxSet <> "" then
		xfout.writeline ct&ct&ct & "& """ & uxSet & """ _"
		pad1st = ", "
	end if
	for each param in refModel.selectNodes("fieldList[masterTable='Y']/field[valueType='input']")
		xpfName = nullText(param.selectSingleNode("fieldName"))
		xfout.writeline ct&ct&ct & "& """ & pad1st & xpfName _
			& "="" & pkStr(request.form(""htx_" & xpfName & """&xn),"""")" & " _"
		pad1st = ", "
	next

	xSelect = refModel.selectSingleNode("sql/selectList").text
	xFrom = refModel.selectSingleNode("sql/fromList").text
	pad1st = " WHERE "
	' -- add primaryKey of masterTable into selectList if it not there
	for each param in refModel.selectNodes("fieldList[masterTable='Y']/field[isPrimaryKey='Y']")
		xpfName = nullText(param.selectSingleNode("fieldName"))
		if instr(xSelect,xpfName&",") = 0 then
			xSelect = "htx." & xpfName & ", " & xSelect
		end if
		xfout.writeline ct&ct&ct & "& """ & pad1st & xpfName _
			& "="" & pkStr(request.form(""xphKey" & xpfName & """&xn),"""")" & " _"
		pad1st = " AND "
	next
	xfout.writeline ct&ct&ct & "& """""
	

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempBList1a.asp"))
	dumpTempFile
    xfin.Close

	
	for each xfk in refModel.selectNodes("fieldList[fkLink]")
		xAlias = xfk.selectSingleNode("fkLink/asAlias").text
		xFrom = "(" & xFrom & " " & xfk.selectSingleNode("fkLink/joinType").text & " JOIN " _
			& xfk.selectSingleNode("tableName").text & " AS " & xAlias & " ON " 
		ckxCount = 0
		for each xfkField in xfk.selectNodes("fkLink/fkFieldList")
			if ckxCount > 0 then	xFrom = xFrom & " AND "
			ckxCount = ckxCount + 1
			xFrom = xFrom & xAlias & "." & xfkField.selectSingleNode("myField").text & " = " _
				& xfk.selectSingleNode("fkLink/refTable").text & "." _
				& xfkField.selectSingleNode("refField").text
		next
		xFrom = xFrom & ")"
	next
	
	xrCount = 0
	for each param in refModel.selectNodes("fieldList/field[refLookup!='']")
	  if nullText(param.selectSingleNode("valueType")) = "direct" then
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeID='" & param.selectSingleNode("refLookup").text & "'"
        SET RSLK=conn.execute(SQL)  
        xAFldName = xAlias & param.selectSingleNode("fieldName").text
        ' --- 把 detailRow 裡的 refField 換掉
        for each xd in xDetail.selectNodes("//colSpec/content/refField[text()='" _
        	& param.selectSingleNode("fieldName").text & "']")
        	xd.text = xAFldName
        next
        ' -----------------------------------
		xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & xAFldName
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("CodeValueFld") & " = " & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("CodeSrcFld")) then _
	    	xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
		xFrom = xFrom & ")"
	  end if
	next

	xfout.writeline ct & "fSql = ""SELECT " & xSelect & cq & " _"
	xfout.writeline ct&ct & "& "" FROM " & xFrom & cq & " _"
	xfout.writeline ct&ct & "& "" WHERE " & refModel.selectSingleNode("sql/whereList").text & cq
	xfout.writeline ct & "xpCondition"
	if nullText(refModel.selectSingleNode("sql/orderList")) <> "" then
		xfout.writeline ct & "fSql = fSql & "" ORDER BY "" & " & cq & refModel.selectSingleNode("sql/orderList").text & cq
	end if

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempBList2.asp"))
	dumpTempFile
    xfin.Close
    
	for each x in htPageDom.selectNodes("//pageSpec/aidLinkList/Anchor")
		ckRight = nullText(x.selectSingleNode("checkRight"))
		if ckRight <> "" then _
			xfout.writeLine ct&ct & cl & "if (HTProgRight and " & ckRight & ")=" & ckRight & " then" & cr
		xfout.writeLine ct&ct&ct & "<A href=""" & x.selectSingleNode("AnchorURI").text _
			& """ title=""" & nullText(x.selectSingleNode("AnchorDesc")) & """>" _
			& x.selectSingleNode("AnchorLabel").text & "</A>"
		if ckRight <> "" then _
			xfout.writeLine ct&ct & cl & "end if" & cr
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempBList20.asp"))
	dumpTempFile
    xfin.Close

	for each x in htPageDom.selectNodes("//pageSpec/taskLinkList/taskLink")
		xfout.writeLine ct&ct & "<option value=""" & nullText(x.selectSingleNode("taskURL")) _
			& """>" & nullText(x.selectSingleNode("taskName")) & "</option>"
	next

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempBList2a.asp"))
	dumpTempFile
    xfin.Close

	for each param in xDetail.selectNodes("colSpec")
		xfout.writeline ct & "<td class=eTableLable>" & nullText(param.selectSingleNode("colLabel")) & "</td>"
	next
    '---------------------------------------
    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempList3.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeLine "<TD align=center><INPUT TYPE=checkbox name=""ckbox" & cl &"=i" & cr _
		& """ value=""Y"">"
	xfout.writeLine cl & "pKey = """"" 
	xpKeyHidden = ""
	for each param in refModel.selectNodes("fieldList[masterTable='Y']/field[isPrimaryKey='Y']")
		xfout.writeLine "pKey = pKey & ""&" & param.selectSingleNode("fieldName").text _
			& "="" & RSreg(""" & param.selectSingleNode("fieldName").text & """)"
		xpkeyHidden = xpKeyHidden & "<INPUT TYPE=hidden name=""xphKey" _
			& param.selectSingleNode("fieldName").text & cl & "=i" & cr & """ value=""" _
			& cl & "=RSreg(""" & param.selectSingleNode("fieldName").text & """)" & cr _
			& """>"
	next	
	xfout.writeLine "if pKey<>"""" then  pKey = mid(pKey,2)" & cr
	xfout.writeLine xpKeyHidden & "</td>"

	for each param in xDetail.selectNodes("colSpec")
		xfout.writeline ct & "<TD class=eTableContent><font size=2>" 
		xUrl = nullText(param.selectSingleNode("url"))
		if  xUrl <> "" then _
			xfout.writeLine ct & "<A href=""" & nullText(param.selectSingleNode("url")) & "?" & cl & "=pKey" & cr & """>"
		processContent param.selectSingleNode("content")
		if xUrl <> "" then	xfout.writeLine "</A>"
		xfout.writeLine "</font></td>"
	next



sub xxxOrgProcessEachColumn
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
end sub
    '---------------------------
    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempList4.asp"))
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

    Set xfin = fs.opentextfile(Server.MapPath("Template0/TempList5.asp"))
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
		processRefField xDom.text
'	  	xfout.writeLine cl & "=RSreg(""" & xDom.text & """)" & cr
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

sub processRefField (xRefField)
	set xrf = refModel.selectSingleNode("fieldList/field[fieldName='" & xRefField & "']")
	if nullText(xrf.selectSingleNode("valueType")) = "direct" then
	  	xfout.writeLine cl & "=RSreg(""" & xRefField & """)" & cr
	else
		processParamField xrf
		xfout.writeLine "<script language=vbs>"
		BListProcessInit xrf
		xfout.writeLine "</script>"
	end if
end sub
%>
</BODY>
</HTML>
