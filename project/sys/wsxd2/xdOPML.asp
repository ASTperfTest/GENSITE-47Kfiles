<%@ CodePage = 65001 %><%
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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\xdOPML.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("mp", "BaseDSD", "CtUnit", "ctNode", "CtNode", "xq_xCat", "xq_xCat2", "debug", "pagesize", "nowPage", "xq_bCat", "xq_bDept")
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
%><% 
dim strOpmlTitle

response.contentType="text/xml" 
strOpmlTitle = "《農業知識入口網》OPML訂閱頻道"

'
%><?xml version="1.0"  encoding="utf-8" ?>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->
<opml>
<head>
    <title><%=strOpmlTitle%></title>
</head>
<body>
<%
dim mp

    mp = request.querystring("mp")

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
	
  	if DSDDom.parseError.reason <> "" then
    		Response.Write("DSDDom parseError on line " &  DSDDom.parseError.line)
    		Response.Write("<BR>Reason: " &  DSDDom.parseError.reason)
    		Response.End()
  	end if
  	
   	set sxRoot = DSDDom.selectSingleNode("DataSchemaDef")
   	set orgDSDRoot = DSDDom.selectSingleNode("DataSchemaDef").cloneNode(true)
	showClientSqlOrderBy = nullText(sxRoot.selectSingleNode("showClientSqlOrderBy"))
	catField = nullText(sxRoot.selectSingleNode("formClientCat"))
	pagesize = nullText(sxRoot.selectSingleNode("pagesize"))
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
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & 1 & ".xml"
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


	sql = "SELECT * FROM CatTreeNode where CtNodeID=" & pkstr(request("ctNode"),"")
	set RS = conn.execute(sql)
	xRootID = RS("CtRootID")
	xCtUnitName = RS("CatName")
	xCtUnit = RS("CtUnitID")
	xPathStr = "<xPathNode Title=""" & deAmp(xCtUnitName) & """ xNode=""" & RS("ctNodeID") & """ />"
	xParent = RS("DataParent")
	mydCondition = RS("dCondition")
	myxslList = RS("xslList")
	'response.write "<xslData>"&myxslList&"</xslData>"

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
	'response.write "<xPath><UnitName>" & deAmp(xCtUnitName) & "</UnitName>" & xpathStr & "</xPath>"

	sqlCom = "SELECT head.xBody AS xBody1, foot.xBody AS xBody2, CtUnitLogo, CtUnitName FROM CtUnit AS C " _
		& " LEFT JOIN CuDtGeneric AS head ON C.headerPart=head.iCuItem" _
		& " LEFT JOIN CuDtGeneric AS foot ON C.footerPart=foot.iCuItem" _
		& " WHERE ctunitid=" & pkStr(xCtUnit,"")
	set RS = conn.execute(sqlcom)
    'response.write sqlcom


	footerPart = RS("xBody2")
	CtUnitLogo = RS("CtUnitLogo")
	CtUnitName = RS("CtUnitName")

          '2006/2/17 頭標改寫  hying

	headerPart = RS("xBody1")
        sqlcom="select * from cudtgeneric a,cattreenode b where a.ictunit=" & pkStr(xCtUnit,"") & " and a.fctupublic = 'N' and a.ximportant='99'  and b.ctnodeid='" & request("ctNode") & "' and  b.dcondition like '%'+a.topcat +'%'"
'response.write sqlcom
        set rs2=conn.Execute(sqlcom)
        if not rs2.eof then
             headerPart = RS2("xBody")
        end if

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
        fSql = fSql	& " AND (ghtx.avEnd is NULL OR ghtx.avEnd >=" & pkStr(date(),")") _
			& " AND (ghtx.avBegin is NULL OR ghtx.avBegin <=" & pkStr(date(),")")
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

	if showClientSqlOrderBy<>"" then
		fSql = fSql & " " & showClientSqlOrderBy
	else
		fSql = fSql & " ORDER BY xPostDate DESC"
	end if


if request("debug")=1 then response.write "<fSql><![CDATA[" & fSql & "]]></fSql>"
'response.end

      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 and pagesize <> "" then
       PerPageSize=pagesize
      elseif  PerPageSize <= 0 then    PerPageSize=1500
      end if
	nowPage=cint(Request.QueryString("nowPage"))  '現在頁數
    if nowPage <= 0 then  nowPage = 1
	totPage=0
	totRec=0

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
	RSreg.CacheSize = PerPageSize
	
'response.write fsql
'response.end
set RSreg = conn.execute (fsql)

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
<%
rem    <outline type="rss" text="" title="" description="" xmlUrl="" htmlUrl="" /> 
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

		doFP

         RSreg.moveNext
         if RSreg.eof then exit for
	next
end if

RSreg.close


sub doFP
    'on error resume Next
	xrCount = 0%>
    <outline type="rss" text="<%=deAmp(RSreg("stitle"))%>"
		title="<%=deAmp(RSreg("stitle"))%>"
		description="<%=deAmp(RSreg("xBody"))%>" 
		xmlUrl="<%=session("myWWWsite")%>/<%=deAmp(RSreg("xURL"))%>"

		<%
		dim RsLpUnit
		dim RsLpBaseDSD

		dim getLpNode
		dim getLpUnit
		dim getLpBaseDSD
		
		getLpNode    = ""
		getLpUnit    = ""
		getLpBaseDSD = ""

		getLpNode=MaxGetField(RSreg("xURL"),"=",1)
		getLpNode=MaxGetField(getLpNode,"&",0)
        'set RsLpUnit = conn.execute("SELECT CtUnitID FROM CatTreeNode WHERE CtNodeID = " & getLpNode & "")
        'if not RsLpUnit.eof then
        '    getLpUnit = trim("" & RsLpUnit("CtUnitID"))

         '   set RsLpBaseDSD = conn.execute("SELECT [ibaseDsd] FROM [CtUnit] where ctunitid = " & getLpUnit & "")
         '   if not RsLpBaseDSD.eof then
         '       getLpBaseDSD = trim("" & RsLpBaseDSD("ibaseDsd"))
         '   end if
        'end if
		%>
		htmlUrl="<%=session("myWWWsite")%>/np.asp?CtNode=<%=getLpNode%>&amp;mp=<%=mp%>"
		
		/>
<%
end sub
%>
</body>
</opml>


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


function MaxGetField(byVal strFullString , byval AM, byval iNumber)
    dim rtnVal
    dim arrName

    rtnVal=""

    if strFullString <> "" and AM <> "" and instr(strFullString,AM) > 0 then
        arrName= split(strFullString,AM)
        rtnVal=arrName(iNumber)
    else
        rtnVal=strFullString
    end if 
    
    MaxGetField = rtnVal
end function

%>