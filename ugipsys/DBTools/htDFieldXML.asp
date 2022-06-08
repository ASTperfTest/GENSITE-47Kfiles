<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCode="HT011"
HTProgPrefix="htDField" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<%

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/HTSDSystem/schema0.xml")
	xv = oxml.load(LoadXML)
  if oxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if
 
	set fxml = server.createObject("microsoft.XMLDOM")
	fxml.async = false
	fxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/HTSDSystem/schemaField0.xml")
	xv = fxml.load(LoadXML)
  if fxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  fxml.parseError.line)
    Response.Write("<BR>Reason: " &  fxml.parseError.reason)
    Response.End()
  end if

  	set dXml = oxml.selectSingleNode("DataSchemaDef/dsTable")


	sqlCom = "SELECT * FROM HtDentity WHERE entityId=" & pkStr(request.queryString("entityID"),"")
	Set RSmaster = Conn.execute(sqlcom)

	dXml.selectSingleNode("tableName").text = RSmaster("tableName")
	dXml.selectSingleNode("tableDesc").text = RSmaster("entityDesc")
	
	fSql = "SELECT htx.*, xref1.mvalue AS xrxdataType FROM (HtDfield AS htx LEFT JOIN CodeMain AS xref1 ON xref1.mcode = htx.xdataType AND xref1.codeMetaId=N'htDdataType')"
	fSql = fSql & " WHERE 1=1"
	fSql = fSql & " AND htx.entityId=" & pkStr(RSmaster("entityId"),"")
	fSql = fSql & " ORDER BY htx.xfieldSeq"
	
	set RSlist = conn.execute(fSql)

  	set flXml = dXml.selectSingleNode("fieldList")


'on error resume next
'		if error.num > 0 then 
'			response.write RSlist("xfieldName") & "<HR>"
'			response.end
'		end if
'	RSlist.moveNext
	while not RSlist.eof	
'		response.write RSlist("xfieldName") & "<BR>"
		set xFieldNode = fxml.selectSingleNode("DataSchemaField/field").cloneNode(true)
		xFieldNode.selectSingleNode("fieldName").text = RSlist("xfieldName")
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
'  response.write oxml.selectSingleNode("//frame[@name='topFrame']/@src").text

	
	
	
	
'	oxml.selectSingleNode("UseCase/Code").text = request("pfx_APcode")
'	oxml.selectSingleNode("UseCase/APCat").text = request("sfx_APCat")
'	oxml.selectSingleNode("UseCase/Version/Date").text = date()
'	oxml.selectSingleNode("UseCase/Version/Author").text = session("userID")
'  response.write oxml.documentElement.xml
	oxml.save(Server.MapPath("/HTSDSystem/xmlSpec/" & RSmaster("tableName") & ".xml"))

  response.end


%>