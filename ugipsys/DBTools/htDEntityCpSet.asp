<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCode="HT011"
HTProgPrefix="htDField" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<%
	xNewDbName = request("htx_xDbName")
	xNewTblName = request("htx_xTableName")

	sqlCom = "SELECT * FROM HtDentity WHERE entityId=" & pkStr(request("htx_entityID"),"")
	Set RSmaster = Conn.execute(sqlcom)
	
	xNewEntity = "HtDentity"
	if xNewDbName <> "" then	xNewEntity = xNewDbName & ".dbo." & xNewEntity
	xNewField = "HtDfield"
	if xNewDbName <> "" then	xNewField = xNewDbName & ".dbo." & xNewField

	sql = "INSERT INTO " & xNewEntity & "(dbId, tableName, entityDesc, entityUri, apcatId) VALUES(" _
		& pkStr(RSmaster("dbId"),",") _
		& pkStr(xNewTblName,",") _
		& pkStr(RSmaster("entityDesc"),",") _
		& pkStr(RSmaster("entityUri"),",") _
		& pkStr(RSmaster("apcatId"),")")
	response.write sql & "<HR>"
	conn.Execute(SQL)  

	sql = "SELECT @@IDENTITY AS DBID"
	set RSx = conn.execute(sql)
	xNewDBID = RSx("dbId")
'	xNewDBID = 1

	fSql = "SELECT * INTO #xTemphtDField FROM HtDfield WHERE entityId=" & pkStr(RSmaster("entityID"),"")
	conn.Execute(fSQL)  
	fSql = "UPDATE #xTemphtDField SET entityId=" & xNewDBID
	conn.Execute(fSQL)  
	fSql = "INSERT INTO " & xNewField & " SELECT * FROM #xTemphtDField"
	conn.Execute(fSQL)  

	fSql = "SELECT * FROM " & xNewField & " WHERE entityId=" & xNewDBID
	set RSlist = conn.execute(fSql)


	while not RSlist.eof	
		response.write RSlist("xfieldName") & "," & RSlist("entityId") & "<BR>"
	  RSlist.moveNext
	wend

%>