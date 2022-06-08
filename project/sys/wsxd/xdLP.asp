<%@ CodePage = 65001 %>
<% response.contentType="text/xml" %><?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<!--#Include virtual = "/inc/MSclient.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->
<% 
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

sub dxpCondition
	if request("xq_xCat") <> "" then
		fSql = fSql & " AND topCat LIKE '%" & request("xq_xCat") & "%'"
		qURL = qURL & "&amp;xq_xCat=" & Server.URLEncode(request("xq_xCat"))
	end if
	if request("xq_xCat2") <> "" then
		fSql = fSql & " AND " & catField2 & " LIKE '%" & request("xq_xCat2") & "%'"
		qURL = qURL & "&amp;xq_xCat2=" & Server.URLEncode(request("xq_xCat2"))
	end if
	if request("xq_bCat") <> "" then
		fSql = fSql & " AND Ghtx.topCat=" & pkStr(request("xq_bCat"),"")
		qURL = qURL & "&amp;xq_xCat=" & Server.URLEncode(request("xq_xCat"))
	end if
	if request("xq_bDept") <> "" then
		fSql = fSql & " AND Ghtx.idept=" & pkStr(request("xq_bDept"),"")
		qURL = qURL & "&amp;xq_xCat=" & Server.URLEncode(request("xq_xCat"))
	end if
	for each param in orgDSDRoot.selectNodes("//fieldList/field[queryListClient!='']")
	  paramKind = nullText(param.selectSingleNode("paramKind"))
	  paramCode = nullText(param.selectSingleNode("fieldName"))
	  paramKindPad = ""
	  if paramKind = "range" then 	paramKindPad = "S"
	  if request("htx_" & paramCode & paramKindPad) <> "" then
		select case paramKind
		  case "range"
			rangeS = request("htx_" & paramCode & "S")
			rangeE = request("htx_" & paramCode & "E")
			if rangeE = "" then	rangeE=rangeS
			whereCondition = replace(paramCode & " BETWEEN '{0}' and '{1}'", _
				"{0}", rangeS)
			whereCondition = replace(whereCondition, "{1}", rangeE)
			qURL = qURL & "&amp;htx_" & paramCode & "S=" & Server.URLEncode(rangeS) _
						& "&amp;htx_" & paramCode & "E=" & Server.URLEncode(rangeE)
		  case "value"
			whereCondition = replace(paramCode & " = {0}", "{0}", _
				pkStr(request("htx_" & paramCode),""))
			qURL = qURL & "&amp;htx_" & paramCode & "=" & Server.URLEncode(request("htx_" & paramCode))
		  case else		'-- LIKE
			whereCondition = replace(paramCode & " LIKE {0}", "{0}", _
				pkStr("%"&request("htx_" & paramCode)&"%",""))
			qURL = qURL & "&amp;htx_" & paramCode & "=" & Server.URLEncode(request("htx_" & paramCode))
		end select
		fSql = fSql & " AND " & whereCondition
'		response.write( "<whereCondition>" & whereCondition & "</whereCondition>")
	  end if
	next
end sub

'-------準備前端呈現需要呈現欄位的DSD xmlDOM
	SQLTable="Select sBaseTableName from BaseDSD where iBaseDSD=" & pkStr(request.queryString("BaseDSD"),"")
	Set RSTable=conn.execute(SQLTable)		
    	'----找出對應的CtUnitX???? xmlSpec檔案(若找不到則抓default), 並依fieldSeq排序成物件存入session
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & request.querystring("CtUnit") & ".xml")
    	if fso.FileExists(filePath) then
    		LoadXMLDSD = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & request.querystring("CtUnit") & ".xml")			
    	else
    		LoadXMLDSD = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & request.querystring("BaseDSD") & ".xml")			
    	end if 
	set DSDDom = Server.CreateObject("MICROSOFT.XMLDOM")
	response.write LoadXMLDSD
	xv = DSDDom.load(LoadXMLDSD)
'	response.write xv & "<HR>"
  	if DSDDom.parseError.reason <> "" then 
    		Response.Write("DSDDom parseError on line " &  DSDDom.parseError.line)
    		Response.Write("<BR>Reason: " &  DSDDom.parseError.reason)
    		Response.End()
  	end if	
   	set sxRoot = DSDDom.selectSingleNode("DataSchemaDef") 	
   	set orgDSDRoot = DSDDom.selectSingleNode("DataSchemaDef").cloneNode(true)
	showClientSqlOrderBy = nullText(sxRoot.selectSingleNode("showClientSqlOrderBy"))
	catField = nullText(sxRoot.selectSingleNode("formClientCat"))
	if catField<>"" then
'		response.write catField & "<HR/>"
		set sxField = sxRoot.selectSingleNode("//field[fieldName='topCat']")
		'response.write catField
		'response.End()
		catCode = nullText(sxField.selectSingleNode("refLookup"))
        if catCode <>"" then
			xSelect = "htx.*, ghtx.*"
			xFrom = nullText(sxRoot.selectSingleNode("//tableName")) & " AS htx " _
				& " JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem "

			SQL="Select * from CodeMetaDef where codeID='" & catCode & "'"
			'response.write SQL
			'response.end
        	SET RSLK=conn.execute(SQL)  
			xSelect = xSelect & ", xref." & RSLK("CodeDisplayFld") & " AS xrefCat"
			xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS xref ON xref." _
				& RSLK("CodeValueFld") & " = " & catField
			if not isNull(RSLK("CodeSrcFld")) then _
	   	 		xFrom = xFrom & " AND xref." & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
			xFrom = xFrom & ")"
'			response.write xFrom & "<HR/>"

			SQL="Select * from CodeMetaDef where codeID='" & catCode & "'"
	       	 SET RSLK=conn.execute(SQL)
			if not RSLK.EOF then
		 	 if isNull(RSLK("CodeSortFld")) then
				if isNull(RSLK("CodeSrcFld")) then
		   	 		Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") 
				else	
		    		Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
				end if          
		 	 else
				if isNull(RSLK("CodeSrcFld")) then
		  		  	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL Order by " & RSLK("CodeSortFld")
		    	else	
		    		Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL AND " & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "' Order by " & RSLK("CodeSortFld") 
				end if
		 	 end if
		 	 Set RSS = conn.execute(Sql)
		 	 response.write "<CatList>"
		 	 while not RSS.eof
		  			response.write "<CatItem><CatCode>" & RSS(0) _
		  				& "</CatCode><CatName>" & RSS(1) _
		  				& "</CatName><xqCondition>&amp;xq_xCat=" & Server.URLEncode(RSS(0)) _
		  				& "</xqCondition></CatItem>"
		  		RSS.moveNext
		 	 wend
				response.write "</CatList>"
			end if	
		end if
	end if
	'丟出第二組資料大類
	catField2 = nullText(sxRoot.selectSingleNode("formClientCat2"))
	if catField2<>"" then
'		response.write catField2 & "<HR/>"
		set sxField = sxRoot.selectSingleNode("//field[fieldName='" & catField2 & "']")
		catCode = nullText(sxField.selectSingleNode("refLookup"))

		xSelect = "htx.*, ghtx.*"
		xFrom = nullText(sxRoot.selectSingleNode("//tableName")) & " AS htx " _
			& " JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem "

		SQL="Select * from CodeMetaDef where codeID='" & catCode & "'"
        	SET RSLK=conn.execute(SQL)  
		xSelect = xSelect & ", xref." & RSLK("CodeDisplayFld") & " AS xrefCat"
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS xref ON xref." _
			& RSLK("CodeValueFld") & " = " & catField2
		if not isNull(RSLK("CodeSrcFld")) then _
	    		xFrom = xFrom & " AND xref." & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
		xFrom = xFrom & ")"
'		response.write xFrom & "<HR/>"

		SQL="Select * from CodeMetaDef where codeID='" & catCode & "'"
	        SET RSLK=conn.execute(SQL)
		if not RSLK.EOF then
		  if isNull(RSLK("CodeSortFld")) then
			if isNull(RSLK("CodeSrcFld")) then
		    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") 
			else	
		    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
			end if          
		  else
			if isNull(RSLK("CodeSrcFld")) then
		    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL Order by " & RSLK("CodeSortFld")
		    else	
		    	Sql = "Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL AND " & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "' Order by " & RSLK("CodeSortFld") 
			end if
		  end if
		  Set RSS = conn.execute(Sql)
		  response.write "<CatList2>"
		  while not RSS.eof
		  		response.write "<CatItem><CatCode>" & RSS(0) _
		  			& "</CatCode><CatName>" & RSS(1) _
		  			& "</CatName><xqCondition>&amp;xq_xCat2=" & Server.URLEncode(RSS(0)) _
		  			& "</xqCondition></CatItem>"
		  	RSS.moveNext
		  wend
			response.write "</CatList2>"
		end if	
	end if

		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
'		LoadXML = server.MapPath("GipDSD") & "\xdmp" & request("mp") & ".xml"
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & request("mp") & ".xml"
'		response.write LoadXML & "<HR>"
		xv = htPageDom.load(LoadXML)
'		response.write xv & "<HR>"
  		if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  		end if

  	set refModel = htPageDom.selectSingleNode("//MpDataSet")
	myTreeNode = request("ctNode")
	response.write "<MenuTitle>"&nullText(refModel.selectSingleNode("MenuTitle"))&"</MenuTitle>"
    response.write "<myStyle>"&nullText(refModel.selectSingleNode("MpStyle"))&"</myStyle>"
	response.write "<mp>"&request("mp")&"</mp>"

	sql = "SELECT * FROM CatTreeNode where CtNodeID=" & pkstr(request("ctNode"),"")
	set RS = conn.execute(sql)
	xRootID = RS("CtRootID")
	xCtUnitName = RS("CatName")
	xCtUnit = RS("CtUnitID")
	xPathStr = "<xPathNode Title=""" & deAmp(xCtUnitName) & """ xNode=""" & RS("ctNodeID") & """ />"
	xParent = RS("DataParent")
	mydCondition = RS("dCondition")
	myxslList = RS("xslList")	
	response.write "<xslData>"&myxslList&"</xslData>"
	%>
	<!--#include file = "gensite.inc" -->
	<!--#include file= "content.inc" -->
	<%

'  if myxslList<>"" then
    	set root = DSDDom.selectSingleNode("DataSchemaDef") 	
    	'----Load XSL樣板
    	set oxsl = server.createObject("microsoft.XMLDOM")
   	oxsl.async = false
   	xv = oxsl.load(server.mappath("/site/" & session("mySiteID") & "/GipDSD/CtUnitXOrder.xsl"))   
    	'----複製Slave的dsTable,並依順序轉換
	set DSDNode = DSDDom.selectSingleNode("DataSchemaDef/dsTable[tableName='"&RSTable(0)&"']").cloneNode(true)    
    	set DSDNodeXML = server.createObject("microsoft.XMLDOM")
   	DSDNodeXML.appendchild DSDNode
    	set nxml = server.createObject("microsoft.XMLDOM")
    	nxml.LoadXML(DSDNodeXML.transformNode(oxsl))
    	set nxmlnewNode = nxml.documentElement  
	for each param in nxmlnewNode.selectNodes("field[showListClient='']") 
		set romoveNode=nxmlnewNode.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		nxmlnewNode.removeChild romoveNode
	next    	  
    	DSDNode.replaceChild nxmlnewNode,DSDNode.selectSingleNode("fieldList")    	
    	root.replaceChild DSDNode,root.selectSingleNode("dsTable[tableName='"&RSTable(0)&"']")
    	'----複製CuDtGeneric的dsTable,並依順序轉換
    	set GenericNode = DSDDom.selectSingleNode("DataSchemaDef/dsTable[tableName='CuDtGeneric']").cloneNode(true)    
    	set GenericNodeXML = server.createObject("microsoft.XMLDOM")
    	GenericNodeXML.appendchild GenericNode
   	set nxml2 = server.createObject("microsoft.XMLDOM")
    	nxml2.LoadXML(GenericNodeXML.transformNode(oxsl))
    	set nxmlnewNode2 = nxml2.documentElement
	for each param in nxmlnewNode2.selectNodes("field[showListClient='' and fieldName!='sTitle']") 
		set romoveNode=nxmlnewNode2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		nxmlnewNode2.removeChild romoveNode
	next        	    
    	GenericNode.replaceChild nxmlnewNode2,GenericNode.selectSingleNode("fieldList")
    	root.replaceChild GenericNode,root.selectSingleNode("dsTable[tableName='CuDtGeneric']")       	
  	set DSDrefModel = DSDDom.selectSingleNode("//dsTable")
  	set DSDallModel = DSDDom.selectSingleNode("//DataSchemaDef")
  	'----混合field順序
	set nxml0 = server.createObject("microsoft.XMLDOM")
	nxml0.LoadXML(DSDDom.transformNode(oxsl))
'  end if

	myParent = xParent
	xLevel = RS("DataLevel") - 1
	if RS("CtNodeKind") <> "C" then
		xLevel = xLevel -1
		myTreeNode = xParent
	end if
	upParent = 0
	
	while xParent<>0
		sql = "SELECT * FROM CatTreeNode where CtNodeID=" & pkstr(xParent,"")
		set RS = conn.execute(sql)
		if RS("DataLevel") = xLevel then	upParent = xParent
		xPathStr = "<xPathNode Title=""" & deAmp(RS("CatName")) & """ xNode=""" & RS("ctNodeID") & """ />" & xPathStr
		xParent = RS("DataParent")
	wend
	response.write "<xPath><UnitName>" & deAmp(xCtUnitName) & "</UnitName>" & xpathStr & "</xPath>"

	sqlCom = "SELECT head.xBody AS xBody1, foot.xBody AS xBody2, CtUnitLogo, CtUnitName FROM CtUnit AS C " _
		& " LEFT JOIN CuDtGeneric AS head ON C.headerPart=head.iCuItem" _
		& " LEFT JOIN CuDtGeneric AS foot ON C.footerPart=foot.iCuItem" _
		& " WHERE CtUnitID=" & pkStr(xCtUnit,"")
	set RS = conn.execute(sqlcom)
	headerPart = RS("xBody1")
	footerPart = RS("xBody2")
	CtUnitLogo = RS("CtUnitLogo")
	CtUnitName = RS("CtUnitName")
%>
	<HeaderPart><![CDATA[<%=RS("xBody1")%>]]></HeaderPart>
	<FooterPart><![CDATA[<%=RS("xBody2")%>]]></FooterPart>
	<CtUnitLogo><%=RS("CtUnitLogo")%></CtUnitLogo>
	<CtUnitName><%=deAmp(RS("CtUnitName"))%></CtUnitName>

<%

	fsql = "SELECT htx.*, ghtx.*, u.CtUnitName, xr1.deptName " _
		& " FROM " & nullText(sxRoot.selectSingleNode("//tableName")) & " AS htx " _
		& " JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem " _
		& " JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit" _
		& " LEFT JOIN dept AS xr1 ON xr1.deptID=ghtx.idept" _
		& " WHERE  ghtx.fCTUPublic='Y' " 
'		& " AND iCtUnit = " & pkstr(xCtUnit,"") 

	xSelect = "htx.*, ghtx.*, u.CtUnitName, xr1.deptName"
	xFrom = nullText(DSDrefModel.selectSingleNode("tableName")) & " AS htx " _
			& " JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem " _
			& " JOIN CtUnit AS u ON u.CtUnitID=ghtx.iCtUnit" _
			& " LEFT JOIN dept AS xr1 ON xr1.deptID=ghtx.idept"
	xrCount = 0
	for each param in DSDrefModel.selectNodes("fieldList/field[refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeID='" & param.selectSingleNode("refLookup").text & "'"
        	SET RSLK=conn.execute(SQL)  
        	xAFldName = "xref" & param.selectSingleNode("fieldName").text
		xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & xAFldName
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("CodeValueFld") & " = htx." & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("CodeSrcFld")) then _
	    		xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
			xFrom = xFrom & ")"
	next	
	for each param in DSDallModel.selectNodes("//dsTable[tableName='CuDtGeneric']/fieldList/field[refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeID='" & param.selectSingleNode("refLookup").text & "'"
        	SET RSLK=conn.execute(SQL)  
        	xAFldName = "xref" & param.selectSingleNode("fieldName").text
		xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & xAFldName
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("CodeValueFld") & " = ghtx." & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("CodeSrcFld")) then _
	    		xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
			xFrom = xFrom & ")"
	next	
	fSql = "SELECT " & xSelect & " FROM " & xFrom 
	fSql = fSql & " WHERE ghtx.fCTUPublic='Y' " 
	if RS("CtUnitName")<>"整體查詢" then
		fSql = fSql & " AND iCtUnit = " & request.querystring("CtUnit")	
	end if

	if mydCondition<>"" then	fsql = fsql & " AND " & mydCondition
'		& " WHERE iBaseDSD in (	4,5) "

	qURL = "CtNode="&request.queryString("CtNode")&"&amp;CtUnit="&request.queryString("CtUnit") _
		& "&amp;BaseDSD=" & request.queryString("BaseDSD") & "&amp;mp=" & request.queryString("mp")
	xqURL = qURL
	if request.queryString("xq_xCat")<>"" then
'		qURL = qURL & "&amp;xq_xCat=" & Server.URLEncode(request.queryString("xq_xCat"))
	end if
	if request.queryString("xq_xCat2")<>"" then
'		xqURL = xqURL & "&amp;xq_xCat2=" & Server.URLEncode(request.queryString("xq_xCat2"))
	end if

	dxpCondition

    fSql = fSql & " ORDER BY ximportant desc, xPostDate DESC, iCUItem"
	'if showClientSqlOrderBy<>"" then
	'	fSql = fSql & " " & showClientSqlOrderBy
	'else
	'	fSql = fSql & " ORDER BY ximportant desc, xPostDate DESC, iCUItem"
	'end if

%>
	<qURL><%=qURL%></qURL>
	<xqURL><%=xqURL%></xqURL>
	<info myTreeNode="<%=myTreeNode%>" upParent="<%=upParent%>" myParent="<%=myParent%>" />
<%  	
if request("debug")=1 then response.write "<fSql><![CDATA[" & fSql & "]]></fSql>"
'response.end

      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then
	  		PerPageSize=20		
      end if 
	nowPage=cint(Request.QueryString("nowPage"))  '現在頁數
    if nowPage <= 0 then  nowPage = 1
	totPage=0
	totRec=0

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
	RSreg.CacheSize = PerPageSize
'RSreg.Open fSql,Conn,3,1
	set RSreg = conn.execute(fSql)
if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
      strSql=server.URLEncode(fSql)
   end if    
end if   
	


%>
	<nowPage><%=nowPage%></nowPage>
	<totPage><%=totPage%></totPage>
	<totRec><%=totRec%></totRec>
	<PerPageSize><%=PerPageSize%></PerPageSize>
	<TopicTitleList>
<%
'	response.write sxRoot.xml
'----欄位title
	for each param in nxml0.selectNodes("//fieldList/field")
%>
		<TopicTitle><%=nullText(param.selectSingleNode("fieldLabel"))%></TopicTitle>
<%
	next
%>
	</TopicTitleList>
	<TopicList xNode="<%=request("ctNode")%>">
<%
'沒資料 秀"建置中"
if RSreg.eof then   

if request("isUserSearch")="Y" then
    cdata = "無查詢資料"
else
    cdata = "建置中"
end if
%>                  
		<Article iCuItem="" newWindow="N" xInDateRange=""  IsPostDate ="N" IsPic ="" addtr = "0">
			<NoNum></NoNum>			
			<Caption><![CDATA[<%=cdata %>]]></Caption>
			<Abstract><![CDATA[]]></Abstract>
			<PostDate></PostDate>
 			<DeptName></DeptName>
			<TopCat></TopCat>
			<xURL></xURL>
		</Article>
<%	
end if
If not RSreg.eof then   
	j=0 '記錄序號
	e=0 '記錄tag
    for i=1 to PerPageSize
    	xURL = "ct.asp?xItem=" & RSreg("iCuItem") & "&amp;ctNode=" & request("ctNode")& "&amp;mp=" & request("mp")
    	if RSreg("ibaseDSD") = 2 then	xURL = deAmp(RSreg("xURL"))
    	if RSreg("ibaseDSD") = 9 then	xURL = deAmp(RSreg("xURL"))
    	if RSreg("showType") = 2 then	xURL = deAmp(RSreg("xURL"))
    	if RSreg("showType") = 3 then	xURL = "public/Data/" & RSreg("fileDownLoad")
		if RS("CtUnitName")="整體查詢" then
	    	xURL = "content.asp?cuItem=" & RSreg("iCuItem")& "&amp;mp=" & request("mp")
		end if
		
		xInDateRange = "Y"
'		if not isNull(RSreg("m011_edate")) AND RSreg("m011_edate")<>"" _
'			AND (xStdDay(RSreg("m011_edate")) < xStdDay(date())) then	xInDateRange="N"
		if not isNull(RSreg("xPostDateEnd")) AND RSreg("xPostDateEnd")<>"" _
			AND (xStdDay(RSreg("xPostDateEnd")) < xStdDay(date())) then	xInDateRange="N"
			
		IsPostDate = "N"
		IsPic = "N"
		for each param in nxmlnewNode2.selectNodes("field[formListClient!='' and fieldName='xpostDate']")			
			IsPostDate = "Y"
		next
		for each param in nxmlnewNode2.selectNodes("field[formListClient!='' and fieldName='ximgFile']")			
			IsPic = "Y"
		next
		j=(nowPage-1)*PerPageSize+i
		
		if e >PerPageSize then e = e mod PerPageSize			
		
		k=int(e/4)
		if myxslList<>"" and myxslList<>"style1" and myxslList<>"style2" then
			doFP
		else
			doCP
		end if
		
		e=e+1

         RSreg.moveNext
         if RSreg.eof then exit for
	next 
end if

%>
	</TopicList>
<%
  for each xDataSet in refModel.selectNodes("DataSet[ContentData='Y']")
	processXDataSet
  next

function message(tempstr)
  outstring = ""
  while len(tempstr) > 0
    pos=instr(tempstr, chr(13)&chr(10))
    if pos = 0 then
      outstring = outstring & tempstr & "<p>"
      tempstr = ""
    else
      outstring = outstring & left(tempstr, pos-1) & "<p>"
      tempstr=mid(tempstr, pos+2)
    end if
  wend
  message = outstring
end function

function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
	deAmp = replace(deAmp,"""","&quot;")
end function

sub doCP
%>                  
		<Article iCuItem="<%=RSreg("iCuItem")%>" newWindow="<%=RSreg("xNewWindow")%>" xInDateRange="<%=xInDateRange%>"  IsPostDate ="<%=IsPostDate%>" IsPic ="<%=IsPic%>" addtr = "<%=k%>">
			<NoNum><%=j %> </NoNum>
			<Caption><![CDATA[<%=RSreg("sTitle")%>]]></Caption>
			<Abstract><![CDATA[<%=RSreg("xabstract")%>]]></Abstract>
			<PostDate><%=RSreg("xPostDate")%></PostDate>
 			<DeptName><%=RSreg("deptName")%></DeptName>
			<TopCat><%=RSreg("TopCat")%></TopCat>
			<xURL><%=xURL%></xURL>
<%		if not isNull(RSreg("xImgFile")) then %>
			<xImgFile>public/Data/<%=RSreg("xImgFile")%></xImgFile>
<%		end if %>

		</Article>
<%

end sub

sub doFP
on error resume Next
	xrCount = 0%>
		<Article iCuItem="<%=RSreg("iCuItem")%>" newWindow="<%=RSreg("xNewWindow")%>" xInDateRange="<%=xInDateRange%>"  IsPostDate ="<%=IsPostDate%>" IsPic ="<%=IsPic%>">
		<ctUnitName><%=deAmp(RSreg("ctUnitName"))%></ctUnitName>
<%	for each param in nxml0.selectNodes("//fieldList/field")
		kf = param.selectSingleNode("fieldName").text
		if nullText(param.selectSingleNode("refLookup")) <> "" _
			AND nullText(param.selectSingleNode("inputType"))<>"refCheckbox" _
			AND nullText(param.selectSingleNode("inputType"))<>"refCheckboxOther" then
			xrCount = xrCount + 1
			kf = "xref" & kf
		end if	    	
%>                  
            		<ArticleField>		
				<fieldName><%=kf%></fieldName>
<%		if kf="stitle" then%>
				<xURL newWindow="<%=RSreg("xNewWindow")%>"><%=xURL%></xURL>
<%		end if%>
				<Value><![CDATA[<%=RSreg(kf)%>]]></Value>
            		</ArticleField>
    <%  next%>
 		</Article>   
<%
end sub
qStr = request.QueryString
if instr(qStr,"&memID") > 0 then
	qStr = mid(qStr, 1, Instr(qStr, "&memID") - 1 )
end if
response.write "<qStr>?site=2&amp;" & deAmp(qStr) & "</qStr>"
'--------------頁尾維護單位---------------------
footer_sql = "select footer_dept, footer_dept_url from cattreeroot left join nodeinfo on cattreeroot.ctrootid = nodeinfo.ctrootid where cattreeroot.pvXdmp ='"& request("mp") &"'"
set footer_rs = conn.Execute(footer_sql)
Response.Write "<footer_dept>" & footer_rs("footer_dept") & "</footer_dept>"
Response.Write "<footer_dept_url>" & footer_rs("footer_dept_url") & "</footer_dept_url>"

'--------------頁尾維護單位-end---------------
%>
<!--#include file="x1Menus.inc" -->
<%
	sql = " SELECT a.ctrootid , a.viewcount AS allview,b.ViewCount AS thisYearView, "
	sql = sql & " c.viewcount AS thisMonthView FROM( "
	sql = sql & "	Select ctRootId, sum(ViewCount) as ViewCount "
	sql = sql & "		 from CounterForSubjectByDate WHERE CtRootId = '" & request("mp") & "' "
	sql = sql & "		 GROUP BY  ctRootId) AS a "
    sql = sql & "     LEFT JOIN (   "
	sql = sql & "	Select ctRootId, sum(ViewCount) as ViewCount "
	sql = sql & "		from CounterForSubjectByDate WHERE CtRootId = '" & request("mp") & "' "
	sql = sql & "		and YEAR(ymd) = YEAR(GETDATE()) GROUP BY  ctRootId ) b ON a.ctRootId = b.ctRootId "
    sql = sql & "   LEFT JOIN ( "
    sql = sql & "   Select ctRootId, sum(ViewCount) as ViewCount "
	sql = sql & "		from CounterForSubjectByDate WHERE CtRootId = '" & request("mp") & "' "
	sql = sql & "		and MONTH(ymd) = MONTH(GETDATE()) GROUP BY  ctRootId )c ON a.ctRootId = c.ctRootId " 
    Set rs = conn.Execute(sql)
	If Not rs.EOF Then
		if Not IsNull(rs("allview"))  then
			countAll = CLng(rs("allview"))
		end If
		if Not IsNull(rs("thisYearView")) then
			countThisYear =CLng(rs("thisYearView"))
		end If
		if Not IsNull(rs("thisMonthView"))then
			countThisMoth = CLng(rs("thisMonthView"))
		Else
			countThisMoth = 1
		end If
	Else
		countAll = 1
		countThisYear = 1
		countThisMoth = 1
	End IF
	Response.Write "<CounterAll>" & countAll & "</CounterAll>"
	Response.Write "<CounterThisYear>" & countThisYear & "</CounterThisYear>"
	Response.Write "<CounterThisMonth>" & countThisMoth & "</CounterThisMonth>"
%>
</hpMain>
