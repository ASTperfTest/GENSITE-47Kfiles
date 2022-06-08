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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\relatedDocument.inc"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("xItem")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "-", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
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
	'---p峹---
	response.write "<relatedDocument>"
	
	'http://10.10.5.79/VSMwcf/VSMws.svc/getRelDoc/{siteID}/{docID}?group={group}&top={top}
		
	Dim siteId : siteId = "1"
	Dim docId : docId = request.querystring("xItem")
	Dim group : group = ""
	Dim top : top = "3"
	
	Dim docURL : docURL = session("relatedDocumentURL")
	docURL = replace(docURL, "{siteID}", siteId)
	docURL = replace(docURL, "{docID}", docId)
	docURL = replace(docURL, "{group}", group)
	docURL = replace(docURL, "{top}", top)
	
	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	

	xv = oxml.load(docURL)

  if oxml.parseError.reason <> "" then 
    response.write "<isConnect>N</isConnect>"
	else
		response.write "<isConnect>Y</isConnect>"
		on error resume next
		
		Dim isExist : isExist = ""		
		isExist = oxml.selectSingleNode("string").text
		if isExist = "" then
			response.write "<haveResult>Y</haveResult>"						
			for each param in oxml.selectNodes("RelatedDocument/Group[@siteID='0']/DocID")			
				response.write "<Article>"
				response.write "<siteId>0</siteId>"
				GetDocumentXml param.text, "0"
				response.write "</Article>"
			next						
			for each param in oxml.selectNodes("RelatedDocument/Group[@siteID='1']/DocID")			
				response.write "<Article>"
				response.write "<siteId>1</siteId>"
				GetDocumentXml param.text, "1"
				response.write "</Article>"
			next
			for each param in oxml.selectNodes("RelatedDocument/Group[@siteID='2']/DocID")			
				response.write "<Article>"
				response.write "<siteId>2</siteId>"
				GetDocumentXml param.text, "2"
				response.write "</Article>"
			next						
			for each param in oxml.selectNodes("RelatedDocument/Group[@siteID='3']/DocID")			
				response.write "<Article>"
				response.write "<siteId>3</siteId>"
				GetDocumentXml param.text, "3"
				response.write "</Article>"
			next			
		else
			response.write "<haveResult>N</haveResult>"
		end if
  end if
	
	response.write "</relatedDocument>"
	
	Sub GetDocumentXml( docid, siteid )
		Dim stitle : stitle = ""
		Dim url : url = ""
		if siteid = "0" then
			sql = "SELECT DISTINCT REPORT.REPORT_ID, REPORT.SUBJECT, CATEGORY.CATEGORY_ID, RESOURCE_RIGHT.ACTOR_INFO_ID " & _
						"FROM REPORT INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID INNER JOIN CATEGORY " & _
						"ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID AND CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID " & _
						"INNER JOIN RESOURCE_RIGHT ON 'report@' + REPORT.REPORT_ID = RESOURCE_RIGHT.RESOURCE_ID " & _
						"WHERE (CAT2RPT.DATA_BASE_ID = 'DB020') AND (RESOURCE_RIGHT.ACTOR_INFO_ID IN ('001', '002', '004')) " & _
						"AND (REPORT.REPORT_ID = '" & docid & "')"
			set rs = KMConn.execute(sql)
			if not rs.eof then
				stitle = "<stitle><![CDATA[" & rs("SUBJECT") & "]]></stitle>"
				url = "<url><![CDATA[/CatTree/CatTreeContent.aspx?ReportId=" & rs("REPORT_ID") & "&DatabaseId=DB020&CategoryId=" & rs("CATEGORY_ID") & "&ActorType=" & rs("ACTOR_INFO_ID") & "]]></url>"			
			else
				stitle = "<stitle></stitle>"
				url = "<url></url>"
			end if
		elseif siteid = "1" or siteid = "2" then		
			sql = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.showType, CuDTGeneric.fileDownLoad, " & _
						"CuDTGeneric.xURL, CatTreeNode.CtNodeID, CatTreeRoot.pvXdmp FROM CuDTGeneric INNER JOIN CtUnit " & _
						"ON CuDTGeneric.iCTUnit = CtUnit.CtUnitID INNER JOIN CatTreeNode ON CtUnit.CtUnitID = CatTreeNode.CtUnitID " & _
						"INNER JOIN CatTreeRoot ON CatTreeNode.CtRootID = CatTreeRoot.CtRootID WHERE CuDTGeneric.iCUItem = " & docid
			set rs = conn.execute(sql)
			if not rs.eof then
				stitle = "<stitle><![CDATA[" & rs("sTitle") & "]]></stitle>"
				if rs("showType") = "1" then
					if siteid = "1" then
						url = "<url><![CDATA[/ct.asp?xItem=" & rs("iCUItem") & "&ctNode=" & rs("CtNodeID") & "&mp=1]]></url>"
					elseif siteid = "2" then
						url = "<url><![CDATA[/subject/ct.asp?xItem=" & rs("iCUItem") & "&ctNode=" & rs("CtNodeID") & "&mp=" & rs("pvXdmp") & "]]></url>"
					end if
				elseif rs("showType") = "2" then
					url = "<url><![CDATA[" & rs("xURL") & "]]></url>"
				elseif rs("showType") = "3" then
					url = "<url><![CDATA[/public/Data/" & rs("fileDownLoad") & "]]></url>"					
				end if
			else
				stitle = "<stitle></stitle>"
				url = "<url></url>"
			end if	
		elseif siteid = "3" then
			sql = "SELECT iCUItem, sTitle, topCat FROM CuDTGeneric WHERE iCUItem = " & docid
			set rs = conn.execute(sql)
			if not rs.eof then
				stitle = "<stitle><![CDATA[" & rs("sTitle") & "]]></stitle>"
				url = "<url><![CDATA[/knowledge/knowledge_cp.aspx?ArticleId=" & rs("iCUItem") & "&ArticleType=A&CategoryId=" & rs("topCat") & "]]></url>"
			else
				stitle = "<stitle></stitle>"
				url = "<url></url>"
			end if				
		end if	
		rs.close
		set rs = nothing			
		response.write stitle
		response.write url
	End Sub
	'---end of p峹---
%>	