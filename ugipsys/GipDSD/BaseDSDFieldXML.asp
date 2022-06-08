<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="單元資料定義"
HTProgCode="GE1T01"
HTProgPrefix="BaseDSD" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/GipDSD/xmlSpec/schema0.xml")
	xv = oxml.load(LoadXML)
  if oxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if
  
	set fxml = server.createObject("microsoft.XMLDOM")
	fxml.async = false
	fxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/GipDSD/xmlSpec/schemaField0.xml")
	xv = fxml.load(LoadXML)
  if fxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  fxml.parseError.line)
    Response.Write("<BR>Reason: " &  fxml.parseError.reason)
    Response.End()
  end if

  	set dXml = oxml.selectSingleNode("DataSchemaDef/dsTable")
	xiBaseDSD = request.queryString("iBaseDSD")
	sqlCom = "SELECT * FROM BaseDSD WHERE iBaseDSD=" & pkStr(xiBaseDSD,"")
	Set RSmaster = Conn.execute(sqlcom)
	formID = "CuDTx" & xiBaseDSD
	if not ISNull(RSmaster("sBaseTableName")) then formID=RSmaster("sBaseTableName")
	
	dXml.selectSingleNode("tableDesc").text = RSmaster("sBaseDSDName")
	dXml.selectSingleNode("tableName").text = formID
	
	fSql = "SELECT htx.*, xref1.mValue AS xrxdataType " _
		& " FROM (BaseDSDField AS htx LEFT JOIN CodeMain AS xref1 ON xref1.mCode = htx.xdataType AND xref1.codeMetaID='htDdataType')"
	fSql = fSql & " WHERE 1=1"
	fSql = fSql & " AND htx.iBaseDSD=" & pkStr(RSmaster("iBaseDSD"),"")
	fSql = fSql & " ORDER BY htx.xFieldSeq"
	set RSlist = conn.execute(fSql)

  	set flXml = dXml.selectSingleNode("fieldList")


	while not RSlist.eof	
'		response.write RSlist("xfieldName") & "<BR>"
		set xFieldNode = fxml.selectSingleNode("DataSchemaField/field").cloneNode(true)
		xFieldNode.selectSingleNode("fieldSeq").text = RSlist("xfieldSeq")
		if not ISNull(RSlist("xfieldName")) then
			xFieldNode.selectSingleNode("fieldName").text = RSlist("xfieldName")
		else
			xFieldNode.selectSingleNode("fieldName").text = formID & "F" & RSlist("iBaseField")
		end if
		
		if not ISNull(RSlist("xfieldName")) then
			paramCode = RSlist("xfieldName")
		else
			paramCode = formID & "F" & RSlist("iBaseField")
		end if
		paramType = RSlist("xrxdataType")
		paramSize=10
		if paramType="varchar" then	paramSize = RSlist("xdataLen")
		if paramType="char" then	paramSize = RSlist("xdataLen")	
		xFieldNode.selectSingleNode("fieldLabelInit").text = RSlist("xfieldLabel")
		xFieldNode.selectSingleNode("fieldLabel").text = RSlist("xfieldLabel")
		if not isNUll(RSlist("xfieldDesc")) then _
			xFieldNode.selectSingleNode("fieldDesc").text = RSlist("xfieldDesc")
		xFieldNode.selectSingleNode("dataType").text = RSlist("xrxdataType")
		if not isNUll(RSlist("xdataLen")) then _		
			xFieldNode.selectSingleNode("dataLen").text = RSlist("xdataLen")
		if not isNUll(RSlist("xinputLen")) then _			
			xFieldNode.selectSingleNode("inputLen").text = RSlist("xinputLen")
		xFieldNode.selectSingleNode("isPrimaryKey").text = RSlist("xisPrimaryKey")
		xFieldNode.selectSingleNode("canNull").text = RSlist("xcanNull")
		xFieldNode.selectSingleNode("identity").text = RSlist("xidentity")
		xFieldNode.selectSingleNode("inputType").text = RSlist("xinputType")
		if not isNUll(RSlist("xclientDefault")) then _			
			xFieldNode.selectSingleNode("clientEdit").text = RSlist("xclientDefault")
		if RSlist("inUse") = "Y" then
			xFieldNode.selectSingleNode("showListClient").text = RSlist("xfieldSeq")
			xFieldNode.selectSingleNode("showList").text = RSlist("xfieldSeq")
			xFieldNode.selectSingleNode("formListClient").text = RSlist("xfieldSeq")
			xFieldNode.selectSingleNode("formList").text = RSlist("xfieldSeq")
			xFieldNode.selectSingleNode("queryListClient").text = RSlist("xfieldSeq")
			xFieldNode.selectSingleNode("queryList").text = RSlist("xfieldSeq")
			xFieldNode.selectSingleNode("paramKind").text = "LIKE"
		end if			
		if RSlist("xcanNull") = "N" then
			xFieldNode.selectSingleNode("formListClient").text = RSlist("xfieldSeq")
			xFieldNode.selectSingleNode("formList").text = RSlist("xfieldSeq")
		end if			
'		if error.num > 0 then 
'			response.write RSlist("xfieldName") & "<BR>"
'			response.end
'		end if
		if not isNUll(RSlist("xrefLookup")) then _			
			xFieldNode.selectSingleNode("refLookup").text = RSlist("xrefLookup")
		if not isNUll(RSlist("xrows")) then _			
			xFieldNode.selectSingleNode("rows").text = RSlist("xrows")
		if not isNUll(RSlist("xcols")) then _			
			xFieldNode.selectSingleNode("cols").text = RSlist("xcols")

		flXml.appendChild xFieldNode
	  RSlist.moveNext
	wend
	oxml.save(Server.MapPath("/GipDSD/xmlSpec/" & RSmaster("sBaseTableName") & ".xml"))

response.write "XML schema產生完成!"
response.end


%>