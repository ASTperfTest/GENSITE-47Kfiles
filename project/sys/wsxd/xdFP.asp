<% response.contentType="text/xml" %><?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->
<% 
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

'---start---加入推薦詞彙---檢查是否開啟---2008/09/10---vincent---
	
	Response.Write "<commendword>" 
	sql = "SELECT codeMetaID, mCode, mValue, mSortValue FROM CodeMain WHERE (codeMetaID = 'commendWord') AND (mCode = 'topic')"
	Set crs = conn.execute(sql)
	if not crs.eof then
		if crs("mValue") = "1" then
			response.write "<isopen>Y</isopen>"
		else
			response.write "<isopen>N</isopen>"
		end if
	else
		response.write "<isopen>N</isopen>"
	end if
	crs.close
	set crs = nothing
	response.write  "<ctnode>" & request.querystring("ctNode") & "</ctnode>"
	Response.Write "</commendword>" 
	
	'---end---加入推薦詞彙---檢查是否開啟---2008/09/10---vincent---


Dim RSKey

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

	if request.queryString("ctNode") = "" then
		sql = "SELECT  t.ctNodeID" _
			& " FROM CuDTGeneric AS htx JOIN CatTreeNode AS t " _
			& " ON t.ctUnitID=htx.iCtUnit" _
			& " WHERE htx.iCuItem=" & pkstr(request("cuItem"),"")
		set RS = conn.execute(sql)
		ctNode = RS(0)
		xItem = request("cuItem")
'		response.write sql
'		response.write ctNode
'		response.end
'		response.redirect "ct.asp?xItem=" & request("cuItem") & "&ctNode=" & RS(0)
	else
		ctNode = request("ctNode")
		xItem = request("xItem")
	end if

	'sql = "UPDATE CuDTGeneric SET ClickCount = ClickCount + 1 WHERE iCUItem = '" & xItem & "' "
	'conn.execute(sql)
	
	myTreeNode = ctNode
	response.write "<MenuTitle>"&nullText(refModel.selectSingleNode("MenuTitle"))&"</MenuTitle>"
	response.write "<myStyle>"&nullText(refModel.selectSingleNode("MpStyle"))&"</myStyle>"
	response.write "<mp>"&request("mp")&"</mp>"

	sql = "SELECT * FROM CatTreeNode where CtNodeID=" & pkstr(ctNode,"")
	set RS = conn.execute(sql)
	xRootID = RS("CtRootID")
	xCtUnitName = RS("CatName")
	xPathStr = "<xPathNode Title=""" & deAmp(xCtUnitName) & """ xNode=""" & RS("ctNodeID") & """ />"
	xParent = RS("DataParent")
	myxslList = RS("xslData")
	response.write "<xslData>"&myxslList&"</xslData>"
	%>
	<!--#include file = "gensite.inc" -->
	<!--#include file= "content.inc" -->
	<%
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

%>
	<info myTreeNode="<%=myTreeNode%>" upParent="<%=upParent%>" myParent="<%=myParent%>" />
<%  	
'-------準備前端呈現需要呈現欄位的DSD xmlDOM
	SQLTable="Select BD.sBaseTableName,CG.iBaseDSD,CG.iCTUnit " & _
		"from CuDtGEneric CG Left Join BaseDSD BD ON CG.iBaseDSD=BD.iBaseDSD " & _
		"where CG.iCUItem=" & pkStr(xItem,"")
	Set RSTable=conn.execute(SQLTable)		
    	'----找出對應的CtUnitX???? xmlSpec檔案(若找不到則抓default), 並依fieldSeq排序成物件存入session
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & RSTable("iCTUnit") & ".xml")
    	if fso.FileExists(filePath) then
    		LoadXMLDSD = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & RSTable("iCTUnit") & ".xml")
    	else
    		LoadXMLDSD = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & RSTable("iBaseDSD") & ".xml")
    	end if  
	set DSDDom = Server.CreateObject("MICROSOFT.XMLDOM")
	DSDDom.async = false
	DSDDom.setProperty("ServerHTTPRequest") = true	
	xv = DSDDom.load(LoadXMLDSD)
'	response.write xv & "<HR>"
  	if DSDDom.parseError.reason <> "" then 
    		Response.Write("DSDDom parseError on line " &  DSDDom.parseError.line)
    		Response.Write("<BR>Reason: " &  DSDDom.parseError.reason)
    		Response.End()
  	end if	
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
	for each param in nxmlnewNode.selectNodes("field[formListClient='']") 
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
	for each param in nxmlnewNode2.selectNodes("field[(formListClient='' and fieldName!='stitle') or inputType='hidden']") 
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
	
	IsPostDate = "N"
	IsPic = "N"
	for each param in nxmlnewNode2.selectNodes("field[formListClient!='' and fieldName='xpostDate']")			
		IsPostDate = "Y"
	next
	for each param in nxmlnewNode2.selectNodes("field[formListClient!='' and fieldName='ximgFile']")			
		IsPic = "Y"
	next
	

	doCP
	
	for each xBlock in DSDallModel.selectNodes("refDataBlock")
		dorefDataGroup xBlock
	next

'	for each xBlock in DSDallModel.selectNodes("refDataBlock")
'		dorefDataBlock xBlock
'	next

  if RS("attachCount") > 0 then
	fSql = "SELECT dhtx.*" _
		& " FROM CuDtAttach AS dhtx" _
		& " WHERE bList='Y'" _
		& " AND dhtx.xiCUItem=" & pkStr(RS("iCUItem"),"") _
		& " ORDER BY dhtx.listSeq"
	set RSlist = conn.execute(fSql)
	response.write "<AttachmentList>" & vbCRLF
	while not RSlist.eof
'		response.write "<Attachment><URL><![CDATA[public/Attachment/" & RSlist("nfileName") _
'			& "]]></URL><Caption><![CDATA[" & RSlist("atitle") & "]]></Caption><Attachkind><![CDATA[" & RSlist("AttachkindA") & "]]></Attachkind><Attachtype><![CDATA[" & RSlist("Attachtype") & "]]></Attachtype></Attachment>"
		response.write "<Attachment><URL><![CDATA[public/Attachment/" & RSlist("nfileName") _
			& "]]></URL><Caption><![CDATA[" & RSlist("atitle") & "]]></Caption></Attachment>"

		RSlist.moveNext
	wend
	response.write "</AttachmentList>"
  end if

  if RS("pageCount") > 0 then
	fSql = "SELECT dhtx.*, n.*" _
		& " FROM CuDTPage AS dhtx" _
		& " JOIN CuDtGeneric AS n ON dhtx.nPageID=n.iCuItem" _
		& " WHERE bList='Y'" _
		& " AND dhtx.xiCUItem=" & pkStr(RS("iCUItem"),"") _
		& " ORDER BY dhtx.listSeq"
	set RSlist = conn.execute(fSql)
	response.write "<ReferenceList>" & vbCRLF
	while not RSlist.eof
		response.write "<Reference iCuItem='" & RSlist("nPageID") & "'><URL>content.asp?CuItem=" & RSlist("nPageID")& "&amp;mp=" & request("mp") _
			& "</URL><Caption><![CDATA[" & RSlist("aTitle") & "]]></Caption>" _
			& "<xImgFile>" & RSlist("xImgFile") & "</xImgFile></Reference>"

		RSlist.moveNext
	wend
	response.write "</ReferenceList>"
  end if
	
  keywordRelation RS("iCUItem")
  if not RSKey.eof then
	response.write "<RelatedList>" & vbCRLF
	while not RSKey.eof
		response.write "<RelatedRef><URL>content.asp?CuItem=" & RSKey("iCUItem") & "&amp;mp=" & request("mp") _
			& "</URL><keyWeights>" & RSKey("KeyWeights") _
			& "</keyWeights><cPostDate>" & d7Date(RSKey("xPostDate")) _
			& "</cPostDate><Caption><![CDATA[" & RSKey("sTitle") & "]]></Caption></RelatedRef>"

		RSKey.moveNext
	wend
	response.write "</RelatedList>"
  end if


  for each xDataSet in refModel.selectNodes("DataSet[ContentData='Y']")
	processXDataSet
  next

function keywordRelation(iCUItem)
    '----參數 CuDTGenric.iCUItem
    '----傳回值 true:產生RSKey recordsets(10筆);false:不能產生RSKey recordsets
    '----RSKey欄位
    '--------iCUItem	相關資料自動編號ID值
    '--------sTitle	相關資料標題
    '--------KeyCounts	相關資料之關鍵字詞與搜尋關鍵字詞相同之個數
    '--------KeyWeights	相關資料之關鍵字詞與搜尋關鍵字詞相同之權重總計
    xiCUItem=iCUItem
    SQLK="Select xKeyword, refID from CuDtGeneric where iCUItem =" & xiCUItem
    set RSK=conn.execute(SQLK)
    if not isNull(RSK("xKeyword")) and RSK("xKeyword") <> "" then
    xRefID = RSK("refID")
    if xRefID="" OR isNull(xRefID)	then	xRefID=xiCUItem
	'----組合where子句所需xKeyword字串
	xKeywordStr = ""
	xKeywordArray = split(RSK("xKeyword"),",")
	for i = 0 to ubound(xKeywordArray)
		'----去除權重符號
		xPos = instr(xKeywordArray(i),"*")
		if xPos <> 0 then
			xStr = Left(trim(xKeywordArray(i)),xPos-1)
		else
			xStr = trim(xKeywordArray(i))
		end if
		xKeywordStr = xKeywordStr + "'" + xStr + "',"
	next
	xKeywordStr = Left(xKeywordStr,Len(xKeywordStr)-1)
	'----產生RSKey recordsets
	SQLKey = "SELECT Top 10 CDTK.iCUItem,CDTG.sTitle,max(CDTG.xPostDate) AS xPostDate,count( CDTK.iCUItem) KeyCounts,sum(weight) KeyWeights " & _
		"FROM CuDTKeyword CDTK join CuDtGeneric CDTG On CDTK.iCUItem=CDTG.iCUItem " & _
		"where CDTK.iCUItem <> " & cStr(xiCUItem) & _
		" and CDTK.iCUItem <> " & xRefID & _
		" and CDTK.xKeyword in (" & xKeywordStr & ") " & _
		" and (CDTG.refID is NULL OR CDTG.refID <> " & cStr(xiCUItem) & ") " & _
		" Group by CDTK.iCUItem,CDTG.sTitle Order by KeyWeights DESC ,KeyCounts DESC, CDTK.iCUItem"	
	set RSKey = conn.execute(SQLKey)
	keywordRelation = true
    else
    set RSKey = conn.execute("SELECT * FROM AP WHERE 1=2")
	keywordRelation = false    
    end if
end function

function message(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	message=""
  	exit function
  elseif instr(1,xs,"<P",0)>0 or instr(1,xs,"<BR",0)>0 or instr(1,xs,"<td",0)>0 then
 	message=xs
  	exit function
  end if
  	xs = replace(xs,vbCRLF&vbCRLF,"<P>")
  	xs = replace(xs,vbCRLF,"<BR/>")
  	message = replace(xs,chr(10),"<BR/>")
end function

function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function

sub doCP
	sql = "SELECT  htx.*, xr1.deptName, u.CtUnitName " _
		& ", (SELECT count(*) FROM CuDtAttach AS dhtx" _
		& " WHERE bList='Y' AND dhtx.xiCUItem=htx.iCuItem) AS attachCount " _
		& ", (SELECT count(*) FROM CuDTPage AS phtx" _
		& " WHERE bList='Y' AND phtx.xiCUItem=htx.iCuItem) AS pageCount " _
		& " FROM CuDtGeneric AS htx LEFT JOIN dept AS xr1 ON xr1.deptID=htx.idept" _
		& " LEFT JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit" _
		& " WHERE htx.iCuItem=" & pkstr(xItem,"")
'	response.write sql
	set RS = conn.execute(sql)
	if not RS.eof then
	cPostDate = d7date(RS("xPostDate"))
	scPostDate = "中華民國 " & mid(cPostDate,1,2) & " 年 " & mid(cPostDate,4,2) & " 月 " & mid(cPostDate,7,2) & " 日" 
%>
		<MainArticle iCuItem="<%=RS("iCuItem")%>" newWindow="<%=RS("xNewWindow")%>" IsPostDate ="<%=IsPostDate%>" IsPic ="<%=IsPic%>">		
			<Caption><%=RS("sTitle")%></Caption>
			<Content><![CDATA[<%=message(RS("xBody"))%>]]></Content>
			<abstract><%=RS("xabstract")%></abstract>
			<PostDate><%=RS("xPostDate")%></PostDate>
			<cPostDate><%=scPostDate%></cPostDate>
 			<DeptName><%=RS("deptName")%></DeptName>
			<TopCat><%=RS("TopCat")%></TopCat>
			<vgroup><%=RS("vgroup")%></vgroup>
			<CtUnitName><%=deAmp(RS("CtUnitName"))%></CtUnitName>
			<xURL><%=deAmp(RS("xURL"))%></xURL>
			<xKeyword><%=deAmp(RS("xKeyword"))%></xKeyword>
<%		if not isNull(RS("xImgFile")) then %>
			<xImgFile>public/Data/<%=RS("xImgFile")%></xImgFile>
<%		end if %>
		</MainArticle>
<%		
	end if
end sub

sub doFP

	xSelect = "D.deptName,htx.*, ghtx.*"
	xFrom = nullText(DSDrefModel.selectSingleNode("tableName")) & " AS htx " _
			& " JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem " _
			& " LEFT JOIN Dept AS D ON D.deptId=ghtx.idept "
	xrCount = 0
	for each param in DSDrefModel.selectNodes("fieldList/field[refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeId='" & param.selectSingleNode("refLookup").text & "'"
		'response.write SQL  & "<HR>"
        	SET RSLK=conn.execute(SQL)  
        	xAFldName = "xref" & param.selectSingleNode("fieldName").text
		xSelect = xSelect & ", " & xAlias & "." & RSLK("codeDisplayFld") & " AS " & xAFldName
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("codeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("codeValueFld") & " = htx." & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("codeSrcFld")) then _
	    		xFrom = xFrom & " AND " & xAlias & "." & RSLK("codeSrcFld") & "='" & RSLK("codeSrcItem") & "'"
			xFrom = xFrom & ")"
	next	
	for each param in DSDallModel.selectNodes("//dsTable[tableName='CuDtGeneric']/fieldList/field[refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeId='" & param.selectSingleNode("refLookup").text & "'"
        	SET RSLK=conn.execute(SQL)  
        	xAFldName = "xref" & param.selectSingleNode("fieldName").text
		xSelect = xSelect & ", " & xAlias & "." & RSLK("codeDisplayFld") & " AS " & xAFldName
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("codeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("codeValueFld") & " = ghtx." & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("codeSrcFld")) then _
	    		xFrom = xFrom & " AND " & xAlias & "." & RSLK("codeSrcFld") & "='" & RSLK("codeSrcItem") & "'"
			xFrom = xFrom & ")"
	next	
	fSql = "SELECT (SELECT count(*) FROM CuDtAttach AS dhtx" _
		& " WHERE bList='Y' AND dhtx.xiCUItem=ghtx.iCuItem) AS attachCount " _
		& ", (SELECT count(*) FROM CuDTPage AS phtx" _
		& " WHERE bList='Y' AND phtx.xiCUItem=ghtx.iCuItem) AS pageCount, " 
	fSql = fSql & xSelect & " FROM " & xFrom 
	fSql = fSql & " WHERE ghtx.iCuItem=" & pkstr(xItem,"")
	fSql = fSql & " ORDER BY xPostDate DESC"

'response.write fSql & "<HR>"
'response.end
	set RS = conn.execute(fSql)

		xInDateRange = "Y"
'		if not isNull(RS("m011_edate")) AND RS("m011_edate")<>"" _
'			AND (xStdDay(RS("m011_edate")) < xStdDay(date())) then	xInDateRange="N"
		if not isNull(RS("xPostDateEnd")) AND RS("xPostDateEnd")<>"" _
			AND (xStdDay(RS("xPostDateEnd")) < xStdDay(date())) then	xInDateRange="N"

	
	if not RS.eof then
	xrCount = 0
%>
		<MainArticle iCuItem="<%=RS("iCuItem")%>" newWindow="<%=RS("xNewWindow")%>" xInDateRange="<%=xInDateRange%>" IsPostDate ="<%=IsPostDate%>" IsPic ="<%=IsPic%>">
<%	for each param in nxml0.selectNodes("//fieldList/field")
		kf = param.selectSingleNode("fieldName").text
		if nullText(param.selectSingleNode("refLookup")) <> "" _
			AND nullText(param.selectSingleNode("inputType"))<>"refCheckbox" _
			AND nullText(param.selectSingleNode("inputType"))<>"refCheckboxOther" then
			xrCount = xrCount + 1
			kf = "xref"  & kf
		elseif nullText(param.selectSingleNode("fieldName"))="idept" then
			kf="deptName"
		end if		
		
%>            		
            		<MainArticleField>	
				<fieldName><%=param.selectSingleNode("fieldName").text%></fieldName>
				<Title><%=param.selectSingleNode("fieldLabel").text%></Title>
				<Value><![CDATA[<%=message(RS(kf))%>]]></Value>
            		</MainArticleField>				
<%
	next
%>		
		</MainArticle>
<%		
	end if
end sub

sub dorefDataGroup(rNode)

	response.write "<refDataBlock>"
	response.write rNode.selectSingleNode("DataLable").xml
	response.write rNode.selectSingleNode("DataRemark").xml

	xSelect = nullText(rNode.selectSingleNode("sqlSelect"))
	xFrom = nullText(rNode.selectSingleNode("sqlFrom"))
	HeaderCount = nullText(rNode.selectSingleNode("SqlTop"))
	if HeaderCount<>"" then HeaderCount = "TOP " & HeaderCount & " "
	xWhere = " WHERE 1=1"
	for each fkFieldRef in rNode.selectNodes("fkFieldRef")
		xWhere = xWhere & " AND " & nullText(fkFieldRef.selectSingleNode("refField")) _
			& "=" & pkstr(RS(nullText(fkFieldRef.selectSingleNode("myField"))),"")
	next
	
	fSql = "SELECT " & HeaderCount & xSelect & " FROM " & xFrom & xWhere
		
	
	if nullText(rNode.selectSingleNode("SqlCondition"))<>"" then	_
		fsql = fsql & " AND " & nullText(rNode.selectSingleNode("SqlCondition"))
	if nullText(rNode.selectSingleNode("SqlOrderBy"))<>"" then
		fSql = fSql & " ORDER BY " & nullText(rNode.selectSingleNode("SqlOrderBy"))
	else
		fSql = fSql & " ORDER BY xPostDate DESC"
	end if
	response.write "<sql><![CDATA[" & fSql & "]]></sql>"


 Set RSreg = conn.execute(fSql)

		for xi = 0 to RSreg.fields.count-1
			if nullText(rNode.selectSingleNode("groupField[text()='"&RSreg.fields(xi).name&"']"))="" then
				response.write "<nF>" & RSreg.fields(xi).name & "</nF>"
			end if
		next
	orgCPKey = ""
	while not RSreg.eof
		cpKey = ""
		for each xf in rNode.selectNodes("groupKey")
			cpKey = cpKey & RSreg(nullText(xf))
		next
		if cpKey <> orgCPKey then
			if orgCPKey <>"" then	response.write "</refGroup>"
			response.write "<refGroup>"
			for each xf in rNode.selectNodes("groupField")
%>                  
            		<groupField>		
				<fieldName><%=nullText(xf)%></fieldName>
				<Value><![CDATA[<%= message(RSreg(nullText(xf))) %>]]></Value>
            		</groupField>
<%  		next
			orgCPKey = cpKey
		end if
%>
		<refData>
<%
		for xi = 0 to RSreg.fields.count-1
			if nullText(rNode.selectSingleNode("groupField[text()='"&RSreg.fields(xi).name&"']"))="" then
%>                  
            		<ArticleField>		
				<fieldName><%=RSreg.fields(xi).name%></fieldName>
				<Value><![CDATA[<%= RSreg.fields(xi) %>]]></Value>
            		</ArticleField>
    <%  	end if 
    	next%>
 		</refData>   
<%
         RSreg.moveNext
	wend

	if orgCPKey <>"" then	response.write "</refGroup>"
	response.write "</refDataBlock>"
end sub

sub dorefDataBlock(rNode)

	response.write "<refDataBlock>"
	response.write rNode.selectSingleNode("DataLable").xml
	response.write rNode.selectSingleNode("DataRemark").xml

	xSelect = nullText(rNode.selectSingleNode("sqlSelect"))
	xFrom = nullText(rNode.selectSingleNode("sqlFrom"))
	HeaderCount = nullText(rNode.selectSingleNode("SqlTop"))
	if HeaderCount<>"" then HeaderCount = "TOP " & HeaderCount & " "
	xWhere = " WHERE 1=1"
	for each fkFieldRef in rNode.selectNodes("fkFieldRef")
		xWhere = xWhere & " AND " & nullText(fkFieldRef.selectSingleNode("refField")) _
			& "=" & pkstr(RS(nullText(fkFieldRef.selectSingleNode("myField"))),"")
	next
	
	fSql = "SELECT " & HeaderCount & xSelect & " FROM " & xFrom & xWhere
		
	
	if nullText(rNode.selectSingleNode("SqlCondition"))<>"" then	_
		fsql = fsql & " AND " & nullText(rNode.selectSingleNode("SqlCondition"))
	if nullText(rNode.selectSingleNode("SqlOrderBy"))<>"" then
		fSql = fSql & " ORDER BY " & nullText(rNode.selectSingleNode("SqlOrderBy"))
	else
		fSql = fSql & " ORDER BY xPostDate DESC"
	end if
	response.write "<sql><![CDATA[" & fSql & "]]></sql>"

PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=60  
      end if 
	nowPage=cint(Request.QueryString("nowPage"))  '現在頁數
    if nowPage <= 0 then  nowPage = 1
	totPage=0
	totRec=0

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,3,1
Set RSreg = Conn.execute(fSql)
	RSreg.CacheSize = PerPageSize
'----------HyWeb GIP DB CONNECTION PATCH----------


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
        <qURL>xItem=<%=xItem%>&amp;ctNode=<%=ctNode%></qURL> 
	<nowPage><%=nowPage%></nowPage>
	<totPage><%=totPage%></totPage>
	<totRec><%=totRec%></totRec>
	<PerPageSize><%=PerPageSize%></PerPageSize>
<%		
	
'	set RSreg = conn.execute(fSql)
If not RSreg.eof then   

    for i=1 to PerPageSize
%>
		<refData>
<%
		for xi = 0 to RSreg.fields.count-1
%>                  
            		<ArticleField>		
				<fieldName><%=RSreg.fields(xi).name%></fieldName>
				<Value><![CDATA[<%= RSreg.fields(xi) %>]]></Value>
            		</ArticleField>
    <%  next%>
 		</refData>   
<%
         RSreg.moveNext
         if RSreg.eof then exit for
	next 
end if


	response.write "</refDataBlock>"
end sub

function nullAttribute(xnode, xname)
on error resume next
	xs = ""
	xs = xnode.getAttributeNode(xname).value
	nullAttribute = xs
end function

sub xdorefDataBlock(rNode)

	response.write "<refDataBlock>"
	response.write rNode.selectSingleNode("DataLable").xml
	response.write rNode.selectSingleNode("DataRemark").xml
	xSelect = ""
	xFrom = ""
	for each xj in rNode.selectNodes("Entity")
		tablename = nullText(xj.selectSingleNode("table"))
		xFrom = xFrom & " " & nullText(xj.selectSingleNode("join")) & " " & tableName
		myAlias = nullText(xj.selectSingleNode("alias"))
		if myAlias="" then	
			myAlias = tablename
		else
			xFrom = xFrom & " AS " & myAlias
		end if
		myAlias = myAlias & "."
		if nullText(xj.selectSingleNode("join"))<>"" then
		end if
		
		for each pkf in xj.selectNodes("pickField")
			myField = nullAttribute(pkf, "orgField") 
			if myField="" then
				xSelect = xSelect & "," & myAlias & nullText(pkf)
			else
				xSelect = xSelect & "," & myAlias & myField & " AS " & nullText(pkf)
			end if
		next
	next

	HeaderCount = nullText(rNode.selectSingleNode("SqlTop"))
	if HeaderCount<>"" then HeaderCount = "TOP " & HeaderCount & " "
	fSql = "SELECT " & HeaderCount & mid(xSelect,2) & " FROM " & xFrom 
	response.write "<sql><![CDATA[" & fSql & "]]></sql>"
	response.write "</refDataBlock>"
end sub

sub xxxx
	for each fkFieldRef in rNode.selectNodes("fkFieldRef")
		fSql = fSql & " AND " & nullText(fkFieldRef.selectSingleNode("refField")) _
			& "=" & pkstr(RS(nullText(fkFieldRef.selectSingleNode("myField"))),"")
	next

	if nullText(rNode.selectSingleNode("SqlCondition"))<>"" then	_
		fsql = fsql & " AND " & nullText(rNode.selectSingleNode("SqlCondition"))
	if nullText(rNode.selectSingleNode("SqlOrderBy"))<>"" then
		fSql = fSql & " ORDER BY " & nullText(rNode.selectSingleNode("SqlOrderBy"))
	else
		fSql = fSql & " ORDER BY xPostDate DESC"
	end if
	response.write "<sql><![CDATA[" & fSql & "]]></sql>"
	
PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=60  
      end if 
	nowPage=cint(Request.QueryString("nowPage"))  '現在頁數
    if nowPage <= 0 then  nowPage = 1
	totPage=0
	totRec=0

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,3,1
Set RSreg = Conn.execute(fSql)
	RSreg.CacheSize = PerPageSize
'----------HyWeb GIP DB CONNECTION PATCH----------


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
        <qURL>xItem=<%=xItem%>&amp;ctNode=<%=ctNode%></qURL> 
	<nowPage><%=nowPage%></nowPage>
	<totPage><%=totPage%></totPage>
	<totRec><%=totRec%></totRec>
	<PerPageSize><%=PerPageSize%></PerPageSize>
<%		
	
'	set RSreg = conn.execute(fSql)
If not RSreg.eof then   

    for i=1 to PerPageSize

    	xURL = "ct.jsp?xItem=" & RSreg("iCuItem") & "&amp;ctNode=" & nullText(rNode.selectSingleNode("DataNode"))
     	if lcase(xBaseTableName) = "adrotate" then	xURL = deAmp(RSreg("xURL"))
    	if RSreg("ibaseDSD") = 2 then	xURL = deAmp(RSreg("xURL"))
	   	if RSreg("showType") = 2 then	xURL = deAmp(RSreg("xURL"))
    	if RSreg("showType") = 3 then	xURL = "public/Data/" & RSreg("fileDownLoad")
%>
		<Article iCuItem="<%=RSreg("iCuItem")%>" newWindow="<%=RSreg("xNewWindow")%>">
			<xURL newWindow="<%=RSreg("xNewWindow")%>"><%=xURL%></xURL>
<%
		xrCount = 0
		for each param in rbDSDDom.selectNodes("//field[showListClient!='']") 
			kf = nullText(param.selectSingleNode("fieldName"))
			if nullText(param.selectSingleNode("refLookup")) <> "" _
				AND nullText(param.selectSingleNode("inputType"))<>"refCheckbox" _
				AND nullText(param.selectSingleNode("inputType"))<>"refCheckboxOther" then
				xrCount = xrCount + 1
				kf = "xref" & xrCount & kf
			end if	    	
%>                  
            		<ArticleField>		
				<fieldName><%=kf%></fieldName>
<%
    kfvalue=RSreg(kf) %>
				<Value><![CDATA[<%= kfvalue %>]]></Value>
            		</ArticleField>
    <%  next%>
 		</Article>   
<%
         RSreg.moveNext
         if RSreg.eof then exit for
	next 
end if
end sub
qStr = request.QueryString
if Instr(qStr, "&memID") > 0 then
	qStr = mid(qStr, 1, Instr(qStr, "&memID") - 1 )
end if
response.write "<qStr>?site=2&amp;" & deAmp(qStr) & "</qStr>"
%>
<!--#include file="x1Menus.inc" -->
</hpMain>
