<% Response.Expires = 0
HTProgCap="DSDXMLList"
HTProgFunc="DSDXMLList"
HTUploadPath="/public/"
HTProgCode="HT011"
HTProgPrefix="bDinoInfo" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
</head>
<BODY>
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

sub xpCondition
	if session("dcondition") <> "" then
		fSql = fSql & " AND " & session("dcondition")
	end if
end sub
	'session("ctNodeId") = request("ctNodeId")
	sql = "SELECT n.dcondition, u.*, u.ibaseDsd, b.sbaseTableName, r.pvXdmp FROM CatTreeNode AS n LEFT JOIN CtUnit AS u ON u.ctUnitId=n.ctUnitId" _
		& " Left Join BaseDsd As b ON u.ibaseDsd=b.ibaseDsd" _
		& " LEFT JOIN CatTreeRoot AS r ON r.ctRootId=n.ctRootId" _
		& " WHERE n.ctNodeId=" & pkstr(session("ctNodeId"),"")
	'response.write sql
	'response.end

	set RS = Conn.execute(sql)
	session("dcondition") = RS("dcondition")	
	session("ctUnitId") = RS("ctUnitId")
	session("ibaseDsd") = RS("ibaseDsd")
	session("fCtUnitOnly") = RS("fCtUnitOnly")
	session("pvXdmp") = RS("pvXdmp")
	
	if isNull(RS("sbaseTableName")) then
		session("sbaseTableName") = "CuDTx" & session("ibaseDsd")
	else
		session("sbaseTableName") = RS("sbaseTableName")
	end if
	session("CheckYN") = RS("CheckYN")	
	if RS("ctUnitKind")<>"U"  and trim(RS("ibaseDsd") & " ")<>"" and not isnull(RS("ctUnitId")) and session("ctUnitId")<>"" then

		'session("CodeID")=request.querystring("iBaseDSD")
		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true		

	   	Set fso = server.CreateObject("Scripting.FileSystemObject")
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtNodeX" & session("ctNodeId") & ".xml")
		if not fso.FileExists(LoadXML) then
			LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(session("CtUnitID")) & ".xml")
			if not fso.FileExists(LoadXML) then
				LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & session("iBaseDSD") & ".xml")
			end if
		end if

		xv = htPageDom.load(LoadXML)

	  	if htPageDom.parseError.reason <> "" then 
	    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
	    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
	    		Response.End()
	  	end if
		set root = htPageDom.selectSingleNode("DataSchemaDef") 	
		'----Load XSL樣板
		set oxsl = server.createObject("microsoft.XMLDOM")
	   	oxsl.async = false
	   	xv = oxsl.load(server.mappath("/GipDSD/xmlspec/CtUnitXOrder.xsl"))    	
	 
		set xDSDNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='"&session("sBaseTableName")&"']")    
		set xDSDNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable")    

		set DSDNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='"&session("sBaseTableName")&"']").cloneNode(true)    
	    set DSDNodeXML = server.createObject("microsoft.XMLDOM")
	   		DSDNodeXML.appendchild DSDNode
	    set nxml = server.createObject("microsoft.XMLDOM")
	    	nxml.LoadXML(DSDNodeXML.transformNode(oxsl))
	    set nxmlnewNode = nxml.documentElement    
	    	DSDNode.replaceChild nxmlnewNode,DSDNode.selectSingleNode("fieldList")
	    	root.replaceChild DSDNode,root.selectSingleNode("dsTable[tableName='"&session("sBaseTableName")&"']")
	    '----複製CuDTGeneric的dsTable,並依順序轉換
	    set GenericNode = htPageDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDtGeneric']").cloneNode(true)    
	    set GenericNodeXML = server.createObject("microsoft.XMLDOM")
	    	GenericNodeXML.appendchild GenericNode
	   	set nxml2 = server.createObject("microsoft.XMLDOM")
	    	nxml2.LoadXML(GenericNodeXML.transformNode(oxsl))
	    set nxmlnewNode2 = nxml2.documentElement    
	    	GenericNode.replaceChild nxmlnewNode2,GenericNode.selectSingleNode("fieldList")
	    	root.replaceChild GenericNode,root.selectSingleNode("dsTable[tableName='CuDtGeneric']")       	

	  '	set session("codeXMLSpec") = htPageDom
	  	'----混合field順序
		set nxml0 = server.createObject("microsoft.XMLDOM")
		nxml0.LoadXML(htPageDom.transformNode(oxsl))
	'	set session("codeXMLSpec2") = nxml0

		'response.write "<XMP>"+htPageDom.xml+"</XMP>" 
		'response.write "<XMP>"+nxml0.xml+"</XMP>" 
		
	  '	set htPageDom = session("codeXMLSpec")
	  '	set allModel2 = session("codeXMLSpec2")  	
		set allModel2 = nxml0	    	
	  	set refModel = htPageDom.selectSingleNode("//dsTable")
	  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")

		xSelect = " htx.*, ghtx.*"
		xFrom = nullText(refModel.selectSingleNode("tableName")) & " AS htx " _
				& " JOIN CuDtGeneric AS ghtx ON ghtx.icuitem=htx.gicuitem "
		xrCount = 0
		for each param in refModel.selectNodes("fieldList/field[showList!='' and refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
			xrCount = xrCount + 1
			xAlias = "xref" & xrCount
			SQL="Select * from CodeMetaDef where codeId=N'" & param.selectSingleNode("refLookup").text & "'"

			SET RSLK=conn.execute(SQL)  
			xAFldName = xAlias & param.selectSingleNode("fieldName").text
			xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & xAFldName
			xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
				& xAlias & "." & RSLK("CodeValueFld") & " = htx." & param.selectSingleNode("fieldName").text
			if not isNull(RSLK("CodeSrcFld")) then _
				xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "'"
			xFrom = xFrom & ")"
		next
		'for each param in allModel.selectNodes("//dsTable[tableName='CuDtGeneric']/fieldList/field[showList!='' and refLookup!='']")
			'response.write param.xml & "<HR>" & vbCRLF
		'next

		for each param in allModel.selectNodes("//dsTable[tableName='CuDtGeneric']/fieldList/field[showList!='' and refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
			xrCount = xrCount + 1
			xAlias = "xref" & xrCount
			if nullText(param.selectSingleNode("fieldName"))="fctupublic" and session("CheckYN")="Y" then
				param.selectSingleNode("refLookup").text = "isPublic3"
				SQL="Select * from CodeMetaDef where codeId=N'" & param.selectSingleNode("refLookup").text & "'"		
				param.selectSingleNode("refLookup").text = "isPublic"	
			else
				SQL="Select * from CodeMetaDef where codeId=N'" & param.selectSingleNode("refLookup").text & "'"		
			end if		
			SET RSLK=conn.execute(SQL)  
			xAFldName = xAlias & param.selectSingleNode("fieldName").text
			xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & xAFldName
			xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
				& xAlias & "." & RSLK("CodeValueFld") & " = ghtx." & param.selectSingleNode("fieldName").text
			if not isNull(RSLK("CodeSrcFld")) then _
				xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "'"
			xFrom = xFrom & ")"
		next

		fSql = " FROM " & xFrom 
		fSql = fSql & " WHERE 2=2 "
		if (HTProgRight AND 64) = 0 then
			fSql = fSql & " AND ghtx.idept LIKE '" & session("deptID") & "%' "
		end if
		if session("fCtUnitOnly")="Y" then	fSql = fSql & " AND ghtx.ictunit=" & session("CtUnitID") & " "
		if nullText(refModel.selectSingleNode("whereList")) <> "" then _
			fSql = fSql & " AND " & refModel.selectSingleNode("whereList").text
		xpCondition
		session("baseSql") = fSql
		session("xSelectSql") = xSelect


		fSql = session("baseSql")
		xSelect = session("xSelectSql")

		if nullText(allModel.selectSingleNode("//showClientSqlOrderBy")) <> "" then
			fSql = fSql & " " & nullText(allModel.selectSingleNode("//showClientSqlOrderBy"))
		else
			fSql = fSql & " ORDER BY ghtx.xpostDate"
		end if
		fSql = "SELECT TOP 100 " & xSelect & fSql

		Set RSreg = Server.CreateObject("ADODB.RecordSet")
		'RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
		
		'response.write "<P>" & fSql	
		'response.end

'----------HyWeb GIP DB CONNECTION PATCH----------
'		RSreg.Open fSql,Conn,3,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------

		

		while not RSreg.eof
			
			session("RSS_method") = "insert"
			session("RSS_iCuItem") = RSreg.fields(0)
			postURL = "/ws/ws_RSSPool.asp"
			Server.Execute (postURL) 
			
			RSreg.moveNext
		wend
	
	end if 
%>
</body>
</html>
