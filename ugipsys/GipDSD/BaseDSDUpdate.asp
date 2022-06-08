<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCode="GE1T01"
HTProgPrefix="BaseDsd" 
' ============= Modified by Chris, 2006/08/24, 修正 UniCode版不一致 ========================'
'		Document: 950822_智庫GIP擴充.doc
'  modified list:
'	在抑止新增刪除設定 (<sBaseDSDXML><addYN>N</addYN></sBaseDSDXML>) Tag 應是sBaseDSDXML，非原程式處理的sBaseDsdXML
'	 改成 sBaseDsdXMLNode.replaceChild newNode0,sBaseDsdXMLNode.selectSingleNode("sBaseDSDXML")
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/GipDSD/schema0.xml")
	xv = oxml.load(LoadXML)
  if oxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if
  
	set fxml = server.createObject("microsoft.XMLDOM")
	fxml.async = false
	fxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/GipDSD/schemaField0.xml")
	xv = fxml.load(LoadXML)
  if fxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  fxml.parseError.line)
    Response.Write("<BR>Reason: " &  fxml.parseError.reason)
    Response.End()
  end if

  	set dXml = oxml.selectSingleNode("DataSchemaDef/dsTable")
	xibaseDsd = request.queryString("ibaseDsd")
	sqlCom = "SELECT * FROM BaseDsd WHERE ibaseDsd=" & pkStr(xibaseDsd,"")
	Set RSmaster = Conn.execute(sqlcom)
	formID = "CuDTx" & xibaseDsd

	dXml.selectSingleNode("tableDesc").text = RSmaster("sBaseDsdName")
	if not ISNull(RSmaster("sBaseTableName")) then 
		dXml.selectSingleNode("tableName").text = RSmaster("sBaseTableName")	
	else	
		dXml.selectSingleNode("tableName").text = formID
	end if

	'----處理BaseDsdfield部分
	fSql = "SELECT htx.*, xref1.mvalue AS xrxdataType " _
		& " FROM (BaseDsdfield AS htx LEFT JOIN CodeMain AS xref1 ON xref1.mcode = htx.xdataType AND xref1.codeMetaId='htDdataType')"
	fSql = fSql & " WHERE 1=1"
	fSql = fSql & " AND htx.ibaseDsd=" & pkStr(RSmaster("ibaseDsd"),"")
	fSql = fSql & " ORDER BY htx.xfieldSeq"
	set RSlist = conn.execute(fSql)

  	set flXml = dXml.selectSingleNode("fieldList")
	
	if not RSlist.eof then
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
			paramcode = RSlist("xfieldName")
		else
			paramcode = formID & "F" & RSlist("iBaseField")
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
	end if
	'----處理CuDTGeneric部分
	set gXml = oxml.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDtGeneric']/fieldList")
	for each xform in request.form
	    if Left(xform,8)="fieldSeq" then
	        xn=mid(xform,9)
		if not (request("ckbox"&xn)<>"" or request("xmlYN"&xn) = "Y") then
			set removenode=gXml.selectSingleNode("field[fieldSeq='"+request("fieldSeq"&xn)+"']")
			gXml.removeChild(removenode)
		end if
	    end if
	next
'	response.write "<XMP>"+gXml.xml+"</XMP>"
'	response.end
	if not isNull(RSmaster("sBaseDsdXML")) then
	    set sBaseDsdXMLNode = oxml.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDtGeneric']")	
	    set nxml0 = server.createObject("microsoft.XMLDOM")
	    nxml0.LoadXML(""+RSmaster("sBaseDsdXML")+"")
	    set newNode0 = nxml0.documentElement	
	    sBaseDsdXMLNode.replaceChild newNode0,sBaseDsdXMLNode.selectSingleNode("sBaseDSDXML")
'	response.write "<XMP>"+sBaseDsdXMLNode.xml+"</XMP>"
'	response.end	
	end if
'	response.write "DONE!"
'	response.end
	'----回存xml
	oxml.save(server.MapPath("/site/" & session("mySiteID") & "/GipDSD/" & formID & ".xml"))

  response.redirect "BaseDsdEditList.asp?phase=edit&ibaseDsd=" & xibaseDsd
%>