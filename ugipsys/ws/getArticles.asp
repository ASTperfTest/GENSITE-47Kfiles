<?xml version="1.0"  encoding="utf-8" ?>
<mofArticleList>
<!--#include virtual = "/inc/dbFunc.inc" -->
<% 
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

sub dxpCondition
	if request("xq_bCat") <> "" then
		fSql = fSql & " AND Ghtx.topCat=" & pkStr(request("xq_bCat"),"")
	end if
	if request("xq_bDept") <> "" then
		fSql = fSql & " AND Ghtx.iDept=" & pkStr(request("xq_bDept"),"")
	end if
	if request("xq_PostDateS") <> "" then
			rangeS = request("xq_PostDateS")
			rangeE = request("xq_PostDateE")
			if rangeE = "" then	rangeE="2050/1/1"
			whereCondition = replace("ghtx.xPostDate BETWEEN '{0}' and '{1}'", "{0}", rangeS)
			whereCondition = replace(whereCondition, "{1}", rangeE)
		fSql = fSql & " AND " & whereCondition
	end if
end sub

	if not isNumeric(request("CtUnit")) then
		response.write "</ArticleList>"
		response.end
	end if
	
	session("ODBCDSN")="Provider=SQLOLEDB;Data Source=10.10.5.128;User ID=hyGIP;Password=hyweb;Initial Catalog=GIPmof"
'	session("ODBCDSN")="Provider=SQLOLEDB;Data Source=210.69.109.16;User ID=hyMOF;Password=mof0530;Initial Catalog=mofgip"
	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------


	sql = "SELECT * FROM CtUNit WHERE CtUnitID=" & request("CtUnit")
	set RS = conn.execute(sql)
	if RS.eof then
		response.write "</ArticleList>"
		response.end
	end if
	
	xBaseDSD = RS("iBaseDSD")
	xUnitName = RS("CtUnitName")

		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/GipDSD") & "\xmlSpec\CuDtx" & xBaseDSD & ".xml"
'		response.write LoadXML & "<HR>"
		xv = htPageDom.load(LoadXML)
'		response.write xv & "<HR>"
  		if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR/>Reasonxx: " &  htPageDom.parseError.reason)
			response.write "</ArticleList>"
    		Response.End()
  		end if

  	set refModel = htPageDom.selectSingleNode("//dsTable")
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")

	sqlCom = "SELECT htx.*, ghtx.*, xtc.mValue AS xtopCat, d.deptName " _
		& " FROM " & nullText(refModel.selectSingleNode("tableName")) _
		& " AS htx JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCUItem "_
		& " LEFT JOIN codeMain AS xtc ON xtc.mCode = ghtx.topCat AND xtc.codeMetaID=N'topDataCat'" _
		& " LEFT JOIN dept AS d ON d.deptID=ghtx.iDept" _
		& " WHERE ghtx.iCtUnit=" & pkStr(request("CtUnit"),"")
'	Set RSreg = Conn.execute(sqlcom)
	
	xBody = "xBody"

	xSelect = "htx.*, ghtx.*"
	xFrom = nullText(refModel.selectSingleNode("tableName")) & " AS htx " _
			& " JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCUItem "
	xrCount = 0
	for each param in refModel.selectNodes("fieldList/field[formList='Y' and refLookup!='']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeID=N'" & param.selectSingleNode("refLookup").text & "'"
        SET RSLK=conn.execute(SQL)  
        xAFldName = xAlias & param.selectSingleNode("fieldName").text
		xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & xAFldName
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("CodeValueFld") & " = htx." & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("CodeSrcFld")) then _
	    	xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "'"
		xFrom = xFrom & ")"
        ' --- 把 detailRow 裡的 refField 換掉
'        	param.selectSingleNode("fieldName").text = xAFldName
        ' -----------------------------------
	next
	for each param in allModel.selectNodes("//dsTable[tableName='CuDTGeneric']/fieldList/field[formList='Y' and refLookup!='']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeID='" & param.selectSingleNode("refLookup").text & "'"
        SET RSLK=conn.execute(SQL)  
        xAFldName = xAlias & param.selectSingleNode("fieldName").text
		xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & xAFldName
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("CodeValueFld") & " = ghtx." & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("CodeSrcFld")) then _
	    	xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "'"
		xFrom = xFrom & ")"
        ' --- 把 detailRow 裡的 refField 換掉
'        	param.selectSingleNode("fieldName").text = xAFldName
        ' -----------------------------------
	next

	fSql = "SELECT " & xSelect & " FROM " & xFrom 
	fSql = fSql & " WHERE ghtx.iCtUnit= " & pkStr(request("CtUnit")," ")
	if session("fCtUnitOnly")="Y" then	fSql = fSql & " AND ghtx.iCtUnit=" & session("CtUnitID") & " "
	if nullText(refModel.selectSingleNode("whereList")) <> "" then _
		fSql = fSql & " AND " & refModel.selectSingleNode("whereList").text

	dxpCondition

	if nullText(refModel.selectSingleNode("orderList")) <> "" then
		fSql = fSql & " ORDER BY " & refModel.selectSingleNode("orderList").text
	else
		fSql = fSql & " ORDER BY xPostDate DESC"
	end if


	set RSreg = conn.execute(fSql)

	rdCount = 0
	while rdCount < 100 AND not RSreg.eof 
		rdCount = rdCount + 1
		response.write "<mofArticle>" & vbCRLF
		kf = "sTitle"
			response.write "<" & kf & "><![CDATA["
				response.write RSreg(kf)
			response.write "]]></" & kf & ">" & vbCRLF

		for each param in allModel.selectNodes("//fieldList/field[formList='Y']")
			kf = param.selectSingleNode("fieldName").text
			response.write "<" & kf & "><![CDATA["
				response.write RSreg(kf)
			response.write "]]></" & kf & ">" & vbCRLF
		next

		xrCount = 0
		for each param in refModel.selectNodes("fieldList/field[formList='Y' and refLookup!='']")
			xrCount = xrCount + 1
			xAlias = "xref" & xrCount
        	kf = xAlias & param.selectSingleNode("fieldName").text
			response.write "<" & kf & "><![CDATA["
				response.write RSreg(kf)
			response.write "]]></" & kf & ">" & vbCRLF
		next
		for each param in allModel.selectNodes("//dsTable[tableName='CuDTGeneric']/fieldList/field[formList='Y' and refLookup!='']")
			xrCount = xrCount + 1
			xAlias = "xref" & xrCount
        	kf = xAlias & param.selectSingleNode("fieldName").text
			response.write "<" & kf & "><![CDATA["
				response.write RSreg(kf)
			response.write "]]></" & kf & ">" & vbCRLF
		next

		response.write "</mofArticle>" & vbCRLF

		RSreg.moveNext
	wend		     	  
%>
</mofArticleList>
