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
progPath="D:\hyweb\GENSITE\project\web\kmsql.asp"

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
	Conn.Open session("ODBCDSN2")
	atype = request("type")
	'response.write kmtype
	sql = "SELECT DISTINCT TOP (10) REPORT.REPORT_ID, REPORT.ONLINE_DATE, REPORT.CLICK_COUNT, REPORT.CREATE_DATE FROM REPORT INNER JOIN "
	sql = sql & "CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID INNER JOIN CATEGORY ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID AND CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID INNER JOIN "
	sql = sql & "RESOURCE_RIGHT ON 'report@' + REPORT.REPORT_ID = RESOURCE_RIGHT.RESOURCE_ID WHERE (CAT2RPT.DATA_BASE_ID = 'DB020') "
	sql = sql & "AND (RESOURCE_RIGHT.ACTOR_INFO_ID = '002') "
	
	if atype = "A" then
    sql = sql & "ORDER BY REPORT.ONLINE_DATE DESC, REPORT.REPORT_ID DESC"
	elseif atype = "B" then
		sql = sql & "ORDER BY REPORT.CLICK_COUNT DESC, REPORT.REPORT_ID DESC"
	elseif atype = "C" then
	else
		response.write "沒有這種分類"
		response.end
	end if
'response.write sql
'response.end
	Dim InSql	
	Set rs = conn.execute(sql)
	if not rs.eof then
		while not rs.eof 
			InSql = InSql & "'" & rs("REPORT_ID") & "',"
			RS.movenext
		wend
	end if		
	InSql = left(InSql, len(InSql) - 1)
	'response.write insql
	'response.write "<hr>"
	sql = "SELECT DISTINCT REPORT.REPORT_ID, CAT2RPT.CATEGORY_ID, REPORT.SUBJECT, REPORT.ONLINE_DATE, REPORT.CLICK_COUNT, REPORT.CREATE_DATE "
	sql = sql & "FROM REPORT INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID INNER JOIN CATEGORY "
	sql = sql & "ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID "
	sql = sql & "WHERE (REPORT.REPORT_ID IN (" & InSql & ")) AND (CAT2RPT.DATA_BASE_ID = 'DB020') "

	if atype = "A" then
    sql = sql & "ORDER BY REPORT.ONLINE_DATE DESC, REPORT.REPORT_ID DESC"
	elseif atype = "B" then
		sql = sql & "ORDER BY REPORT.CLICK_COUNT DESC, REPORT.REPORT_ID DESC"
	elseif atype = "C" then
	end if		
	'response.write sql
	'response.end
	Set rs2 = conn.execute(sql)
	Dim OldId
	if not rs2.eof then
		
		while not rs2.eof 
			
			If OldId <> rs2("REPORT_ID") Then
				
				link = "<a href=""/CatTree/CatTreeContent.aspx?ReportId=" & rs2("REPORT_ID") & "&DatabaseId=DB020&CategoryId=" & rs2("CATEGORY_ID") & "&ActorType=002"" title=""" & rs2("SUBJECT")& """>" & rs2("SUBJECT")& "</a><span class=""date"">" & rs2("ONLINE_DATE") & "</span>"
				if atype = "A" then
					link = link 
				elseif atype = "B" then
					link = link & " (" & rs2("CLICK_COUNT") & ")"
				end if
				response.write "<li>"& link & "</li>"  & VBCRLF
			End If
			OldId = rs2("REPORT_ID")
			rs2.movenext
		wend
		
	end if	
%>
