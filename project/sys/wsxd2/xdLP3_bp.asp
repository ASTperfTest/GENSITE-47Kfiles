<%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/10 上午 09:59:37
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\xdLP3_bp.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("xq_xCat", "xq_xCat2", "xq_bCat", "xq_bDept", "mp", "CtNode", "data_base_id", "ctNode", "CtUnit", "BaseDSD", "KMautoID", "category_id", "KMAutoID", "memID", "gstyle", "debug", "pagesize", "nowPage")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
genParamsMessage=now() & vbTab & "Error(1):REQUEST變數含特殊字元"

'-------- (可修改)只檢查單引號，如 Request 變數未來將要放入資料庫，請一定要設定(防 SQL Injection) --------
sqlInjParamsArray=array()
sqlInjParamsPattern=array("'")	'#### 要檢查的 pattern(程式會自動更新):DB ####
sqlInjParamsMessage=now() & vbTab & "Error(2):REQUEST變數含單引號"

'-------- (可修改)只檢查 HTML標籤符號，如 Request 變數未來將要輸出，請設定 (防 Cross Site Scripting)--------
xssParamsArray=array()
xssParamsPattern=array("<", ">", "%3C", "%3E")	'#### 要檢查的 pattern(程式會自動更新):HTML ####
xssParamsMessage=now() & vbTab & "Error(3):REQUEST變數含HTML標籤符號"

'-------- (可修改)檢查數字格式 --------
chkNumericArray=array()
chkNumericMessage=now() & vbTab & "Error(4):REQUEST變數不為數字格式"

'-------- (可修改)檢察日期格式 --------
chkDateArray=array()
chkDateMessage=now() & vbTab & "Error(5):REQUEST變數不為日期格式"

'##########################################
ChkPattern genParamsArray, genParamsPattern, genParamsMessage
ChkPattern sqlInjParamsArray, sqlInjParamsPattern, sqlInjParamsMessage
ChkPattern xssParamsArray, xssParamsPattern, xssParamsMessage
ChkNumeric chkNumericArray, chkNumericMessage
ChkDate chkDateArray, chkDateMessage
'--------- 檢查 request 變數名稱 --------
Sub ChkPattern(pArray, patern, message)
	for each str in pArray	
		p=request(str)
		for each ptn in patern
			if (Instr(p, ptn) >0) then
				message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
				Log4U(message) '寫入到 log
				OnErrorAction
			end if
		next
	next
End Sub

'-------- 檢查數字格式 --------
Sub ChkNumeric(pArray, message)
	for each str in pArray
		p=request(str)
		if not isNumeric(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'--------檢察日期格式 --------
Sub ChkDate(pArray, message)
	for each str in pArray
		p=request(str)
		if not IsDate(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'onError
Sub OnErrorAction()
	if (onErrorPath<>"") then response.redirect(onErrorPath)
	response.end
End Sub

'Log 放在網站根目錄下的 /Logs，檔名： YYYYMMDD_log4U.txt
Function Log4U(strLog)
	if (activeLog4U) then
		fldr=Server.mapPath("/") & "/Logs"
		filename=Year(Date()) & Right("0"&Month(Date()), 2) & Right("0"&Day(Date()),2)
		
		filename = filename & "_log4U.txt"
		
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'產生新的目錄
		If (Not fso.FolderExists(fldr)) Then
			Set f = fso.CreateFolder(fldr)
		Else
			ShowAbsolutePath = fso.GetAbsolutePathName(fldr)
		End If
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		'開啟檔案
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile( fldr & "\" & filename , ForAppending, True, -1)
		f.Write strLog  & vbCrLf
	end if
End Function
'##### 結束：此段程式碼為自動產生，註解部份請勿刪除 #####
%><% response.contentType="text/xml" %>
<?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<!--#Include virtual = "/inc/MSclient.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->
<!--#Include file = "time.inc" -->
<% 
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

sub dxpCondition
	if request("xq_xCat") <> "" then
		fSql = fSql & " AND " & catField & " LIKE '%" & request("xq_xCat") & "%'"
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


''----知識樹處理---
'KMDB = "DB001"
''actor_info_id = "002"
'if request("mp") = "3" then
'	KMDB = "DB007"
'	'KMDB = "DB001"
'	'actor_info_id = "004"
'elseif request("mp") = "2" then
'	KMDB = "DB003"
'	'KMDB = "DB001"
'	'actor_info_id = "001"
'else
'	KMDB = "DB004"
'	'KMDB = "DB001"
'	'actor_info_id = "002"
'end if
''response.write KMDB	'DB007
''response.end
'KMautoID = ""
'KMLikeStr = ""	
'xRSSNodeType = ""
'xRSSURL = ""
'SQLNode = "Select CTN.*,R.RSSURL from CatTreeNode CTN Left Join RSSURL R ON CTN.RSSURLID = R.RSSURLID " & _
'					"where CTN.CtNodeID = " & request.queryString("CtNode")
'SET RSNode = conn.execute(SQLNode)
'if not isNull(RSNode("RSSNodeType")) then xRSSNodeType = RSNode("RSSNodeType")
'if not isNull(RSNode("RSSURL")) then xRSSURL = RSNode("RSSURL")
'if not isNull(RSNode("KMautoID")) then KMautoID = RSNode("KMautoID")
''response.write xRSSNodeType
''response.end

'if xRSSNodeType = "1" or xRSSNodeType = "5" or (xRSSNodeType = "6" and request.querystring("data_base_id") <> "") then	
'
'	'---這一段去kmintr抓tree的xml回來---把去抓XML的部份改寫成去LOAD新寫的.asp---
'	'----知識樹RSS URL
'	'---new---
'	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
'	htPageDom.async = false
'	htPageDom.setProperty("ServerHTTPRequest") = true
'	
'	xRSSURL = "http://sgensite/site/coa/wsxd2/xdCatTree.aspx"
'	LoadXML = xRSSURL & "?DatabaseId=" & KMDB & "&CtNode=" & request.querystring("ctNode") & _
'					"&CtUnit=" & request.querystring("CtUnit") & "&BaseDSD=" & request.querystring("BaseDSD")					
'	
'	if request.querystring("KMautoID") <> "" and request.querystring("KMautoID") <> "0" then 
'		
'		LoadXML = xRSSURL & "?DatabaseId=" & request.querystring("data_base_id") & _
'						"&CategoryId=" & request.querystring("category_id") & "&CtNode=" & request.querystring("ctNode") & _
'						"&CtUnit=" & request.querystring("CtUnit") & "&BaseDSD=" & request.querystring("BaseDSD") & _
'						"&KMautoID=" & request.querystring("KMAutoID")
'	end If
'	'---
'	'response.write 	LoadXML
'	'response.end
'	xv = htPageDom.load(LoadXML)
'	if htPageDom.parseError.reason <> "" then 
'		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
'		Response.Write("<BR>Reasonxx1: " &  htPageDom.parseError.reason)
'		Response.End()
'	end if
'	set treeModel = htPageDom.selectSingleNode("//TreeScriptCode")
'	response.write treeModel.xml
'end if



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
		set sxField = sxRoot.selectSingleNode("//field[fieldName='" & catField & "']")
		catCode = nullText(sxField.selectSingleNode("refLookup"))

		xSelect = "htx.*, ghtx.*"
		xFrom = nullText(sxRoot.selectSingleNode("//tableName")) & " AS htx " _
			& " JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem "

		SQL="Select * from CodeMetaDef where codeID='" & catCode & "'"
        	SET RSLK=conn.execute(SQL)  
		xSelect = xSelect & ", xref." & RSLK("CodeDisplayFld") & " AS xrefCat"
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS xref ON xref." _
			& RSLK("CodeValueFld") & " = " & catField
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
	
		response.write "<login>"
		Dim memName
		if request("memID") <> "" Then		
			sql = "SELECT realname FROM Member where account = '" & request("memID") & "'"
			Set memrs = conn.Execute(sql)
			If Not memrs.EOF Then
				response.write "<status>true</status>"	
				memName = memrs("realname")	
			else 
				response.write "<status>false</status>"	
			End If
		else
			response.write "<status>false</status>"
		End If
		response.write "<memID>" & request("memID") & "</memID>"
		response.write "<memName><![CDATA[" & memName & "]]></memName>"
		response.write "<gstyle>" & request("gstyle") & "</gstyle>"	
	response.write "</login>"

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
	if nullText(DSDrefModel.selectSingleNode("whereList")) <> "" then _
		fSql = fSql & " AND " & DSDrefModel.selectSingleNode("whereList").text		
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

	if showClientSqlOrderBy<>"" then
		fSql = fSql & " " & showClientSqlOrderBy
	else
		fSql = fSql & " ORDER BY xPostDate DESC"
	end if

%>
	<qURL><%=qURL%></qURL>
	<xqURL><%=xqURL%></xqURL>
	<info myTreeNode="<%=myTreeNode%>" upParent="<%=upParent%>" myParent="<%=myParent%>" />
<%  	
if request("debug")=1 then response.write "<fSql><![CDATA[" & fSql & "]]></fSql>"
'response.end

      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=15  
      end if 
	nowPage=cint(Request.QueryString("nowPage"))  '現在頁數
    if nowPage <= 0 then  nowPage = 1
	totPage=0
	totRec=0

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
	RSreg.CacheSize = PerPageSize
RSreg.Open fSql,Conn,3,1

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
If not RSreg.eof then   

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

		if myxslList<>"" then
			doFP
		else
			doCP
		end if

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
end function

sub doCP
%>                  
		<Article iCuItem="<%=RSreg("iCuItem")%>" newWindow="<%=RSreg("xNewWindow")%>" xInDateRange="<%=xInDateRange%>">
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
		<Article iCuItem="<%=RSreg("iCuItem")%>" newWindow="<%=RSreg("xNewWindow")%>" xInDateRange="<%=xInDateRange%>">
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
%>

<%
'xGame = nullText(sxRoot.selectSingleNode("//GamePrize/PrizeRate[0]/Prize"))
      PerPageSize_Game=cint(Request.QueryString("pagesize"))
      if PerPageSize_Game <= 0 then  
         PerPageSize_Game=15  
      end if 
	nowPage_Game=cint(Request.QueryString("nowPage"))  '現在頁數
    if nowPage_Game <= 0 then  nowPage_Game = 1
	totPage_Game=0
	totRec_Game=0

 Set rsw = Server.CreateObject("ADODB.RecordSet")
sc="select realname,money from (SELECT name,money, (RANK() OVER (PARTITION BY name ORDER BY money DESC )) as rno FROM flashgame  )as a, member as b where a.rno=1 and a.name=b.account order by money desc"
'set rsw = Conn.execute(sc)
rsw.Open sc,Conn,3,1

if Not rsw.eof then 
   totRec_Game=rsw.Recordcount       '總筆數
   if totRec_Game>0 then 
      
      rsw.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage_Game)<1 then 
         nowPage_Game=1
      elseif cint(nowPage_Game) > rsw.PageCount then 
         nowPage_Game=rsw.PageCount 
      end if            	

      rsw.AbsolutePage=nowPage_Game
      totPage_Game=rsw.PageCount       '總頁數

   end if    
end if 

response.Write "<PerPageSize_Game>" & PerPageSize_Game & "</PerPageSize_Game>"
response.Write "<nowPage_Game>" & nowPage_Game & "</nowPage_Game>"
response.Write "<totPage_Game>" & totPage_Game & "</totPage_Game>"
response.Write "<totRec_Game>" & totRec_Game & "</totRec_Game>"


Response.Write "<coaGame>"		
Do While Not rsw.EOF
	response.Write "<Rank>"		
	response.Write "<Player>" & rsw("realname") & "</Player>"
	response.Write "<Score>" & rsw("money") & "</Score>"
	response.Write "</Rank>"	
  TotPlayer=TotPlayer+1		
rsw.MoveNext				
Loop
response.Write "<TotPlayer>" & TotPlayer & "</TotPlayer>"
Response.Write "</coaGame>"	

Response.Write "<GamePrize>"
sc="select count(*) as Prize from (SELECT name,money, (RANK() OVER (PARTITION BY name ORDER BY money DESC )) as rno FROM flashgame  )as a where a.rno=1 and a.money >5000"
set rsm = Conn.execute(sc)
Do While Not rsm.EOF
	response.Write "<PrizeRate>"		
	response.Write "<Prize>" & cint(rsm("Prize"))/2*100 & "</Prize>"
	response.Write "</PrizeRate>"			
rsm.MoveNext				
Loop

sc="select count(*) as Prize from (SELECT name,money, (RANK() OVER (PARTITION BY name ORDER BY money DESC )) as rno FROM flashgame  )as a where a.rno=1  and a.money>3000 and a.money<=5000"
set rsm = Conn.execute(sc)
Do While Not rsm.EOF
	response.Write "<PrizeRate>"		
	response.Write "<Prize>" & cint(rsm("Prize"))/5*100 & "</Prize>"
	response.Write "</PrizeRate>"			
rsm.MoveNext				
Loop

sc="select count(*) as Prize from (SELECT name,money, (RANK() OVER (PARTITION BY name ORDER BY money DESC )) as rno FROM flashgame  )as a where a.rno=1 and a.money>2000 and a.money<=3000"
set rsm = Conn.execute(sc)
Do While Not rsm.EOF
	response.Write "<PrizeRate>"		
	response.Write "<Prize>" & cint(rsm("Prize"))/25*100 & "</Prize>"
	response.Write "</PrizeRate>"			
rsm.MoveNext				
Loop

sc="select count(*) as Prize from (SELECT name,money, (RANK() OVER (PARTITION BY name ORDER BY money DESC )) as rno FROM flashgame  )as a where a.rno=1 "
set rsm = Conn.execute(sc)
Do While Not rsm.EOF
	response.Write "<PrizeRate>"		
	response.Write "<Prize>" & cint(rsm("Prize"))/50*100 & "</Prize>"
	response.Write "</PrizeRate>"			
rsm.MoveNext				
Loop
Response.Write "</GamePrize>"	
%>
<!--#include file="x1Menus.inc" -->
<%	
	sql = "SELECT * FROM counter where mp='" & request("mp") & "'"
	Set rs = conn.Execute(sql)
	If Not rs.EOF Then
		count = rs("counts") + 1
		sql = "UPDATE counter SET counts = counts + 1  WHERE mp='" & request("mp") & "'"
	Else
		count = 1
		sql="INSERT INTO counter (mp, counts) VALUES ('" & request("mp") & "','1')"
	End If
	Response.Write "<Counter>" & count & "</Counter>"
	Set rs = conn.Execute(sql)
	
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	sql = "SELECT max(xpostDate) FROM CuDtGeneric AS htx JOIN CatTreeNode AS n ON n.CtUnitID=htx.iCtUnit" _
		& " AND n.CtRootID=" & xRootID
	set rs = conn.execute(sql)
%>
<lastupdate><% =year(rs(0)) & "/" & month(rs(0)) &"/" & day(rs(0)) %></lastupdate>
<etoday><%=date()%></etoday>
<today><% ="民國"& year(date())-1911 & "年" & month(date()) &"月"& day(date()) & "日"%></today>
</hpMain>
