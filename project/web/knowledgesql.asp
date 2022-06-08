<%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/9 上午 10:55:19
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\web\knowledgesql.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("type")
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
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open session("ODBCDSN")
	atype = request("type")
	
	if atype = "KA" then
		
    sql = "SELECT TOP (5) iCUItem, topCat, xPostDate, sTitle, CodeMain.mValue FROM CuDTGeneric " & _
    			"INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode INNER JOIN KnowledgeForum " & _
    			"ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem WHERE (iCTUnit = '" & session("KnowledgeQuestionCtUnitId") & "') " & _
    			"AND (CodeMain.codeMetaID = 'KnowledgeType') AND (CuDTGeneric.fCTUPublic = 'Y') " & _
    			"AND (CuDTGeneric.siteId = '" & session("KnowledgeSiteId") & "') AND (KnowledgeForum.Status = 'N') " & _
    			"ORDER BY xPostDate DESC"
    			
	elseif atype = "KB" then
		
		sql = "SELECT TOP (5) CuDTGeneric.iCUItem, CuDTGeneric.topCat, CuDTGeneric.xPostDate, CuDTGeneric.sTitle, " & _
					"CodeMain.mValue FROM CuDTGeneric INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode " & _
					"INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem WHERE (iCTUnit = '" & session("KnowledgeQuestionCtUnitId") & "') " & _
					"AND (CodeMain.codeMetaID = 'KnowledgeType') AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = '" & session("KnowledgeSiteId") & "') " & _
          "AND (KnowledgeForum.Status = 'N') ORDER BY KnowledgeForum.DiscussCount DESC"
          
	elseif atype = "KC" then
		
		sql = "SELECT TOP (5) CuDTGeneric.iCUItem, CuDTGeneric.topCat, CuDTGeneric.xPostDate, CuDTGeneric.sTitle, CodeMain.mValue " & _
					"FROM CuDTGeneric INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode INNER JOIN KnowledgeForum " & _
					"ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem WHERE (KnowledgeForum.Status = 'N') AND (CuDTGeneric.iCTUnit = '" & session("KnowledgeQuestionCtUnitId") & "') " & _
					"AND (CodeMain.codeMetaID = 'KnowledgeType') AND (KnowledgeForum.HavePros = 'Y') AND (CuDTGeneric.fCTUPublic = 'Y') " & _ 
					"AND (CuDTGeneric.siteId = '" & session("KnowledgeSiteId") & "') ORDER BY CuDTGeneric.xPostDate DESC"
		
	else
		response.write "沒有這種分類"
		response.end
	end if
		
	Set rs2 = conn.execute(sql)
	
	if not rs2.eof then
		while not rs2.eof 
		
		 link = "<a href=""/knowledge/knowledge_cp.aspx?ArticleId=" & rs2("iCUItem") & "&ArticleType=A&CategoryId=" & rs2("topCat") & """>" & _
     				rs2("sTitle") & "</a><span class=""from"">" & rs2("mValue") & "</span><span class=""date"">" & DateValue(rs2("xPostDate")) & "</span>"
					
			response.write "<li>"& link & "</li>"  & VBCRLF
			rs2.movenext
		wend	
	else
		response.write "沒有資料"	
	end if	
%>
