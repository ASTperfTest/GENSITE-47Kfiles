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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\time.inc"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("memID")
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
'---
Dim NowTime
NowTime = Now
NowTime = CInt(hour(NowTime))
if NowTime >= 6 And NowTime < 12 Then
	response.write "<NowTime>AM</NowTime>"
elseif NowTime >= 12 And NowTime < 18 Then
	response.write "<NowTime>PM</NowTime>"
else
	response.write "<NowTime>NT</NowTime>"
end if

'---current url---
if instr(request.querystring, "memID") > 2 then
	redircetUrl = left(request.querystring, instr(request.querystring, "memID") - 2)
end if
redircetUrl = replace(redircetUrl, "&", ";")
if instr(Request.ServerVariables("URL"), "xdcp2.asp") > 0 then
	response.write "<redirectUrl><![CDATA[/ct.asp?" & redircetUrl & "]]></redirectUrl>"
elseif instr(Request.ServerVariables("URL"), "xdlp.asp") > 0 then
	response.write "<redirectUrl><![CDATA[/lp.asp?" & redircetUrl & "]]></redirectUrl>"
elseif instr(Request.ServerVariables("URL"), "xdmp.asp") > 0 then
	response.write "<redirectUrl><![CDATA[/mp.asp?" & redircetUrl & "]]></redirectUrl>"
end if
'---

'---activity---
response.write "<activity>"
'---Is activity in the time range---
sql = "SELECT * FROM Activity WHERE (GETDATE() BETWEEN ActivityStartTime AND ActivityEndTime) AND ActivityId = '" & session("ActivityId") & "'"
Set actrs = conn.Execute(sql)
If Not actrs.Eof Then
	response.write "<active>true</active>"
Else
	response.write "<active>false</active>"
End If
actrs.close
set actrs = nothing

'---get the user activity grade---
if request("memID") <> "" Then	
	sql = "SELECT (DiscussCheckGrade + QuestionCheckGrade) AS Grade FROM ActivityMemberNew WHERE MemberId = '" & request("memID") & "'"
	Set actmemrs = conn.Execute(sql)
	If Not actmemrs.Eof Then
		response.write "<ActivityGrade>" & actmemrs("Grade") & "</ActivityGrade>"
	Else
		response.write "<ActivityGrade>0</ActivityGrade>"
	End If
	actmemrs.close
	Set actmemrs = Nothing
End If
response.write "</activity>"

	response.write "<webActivity>"

	sql = "SELECT * FROM m011 WHERE GETDATE() BETWEEN m011_bdate AND m011_edate AND m011_online = '1'"
	set rs = conn.execute(sql)
	If rs.eof Then 
		response.write "<active>false</active>"
		response.write "<id></id>"
	else
		response.write "<active>true</active>"
		response.write "<id>" & rs ("m011_subjectid") & "</id>"
	end if
	rs.close
	set rs = nothing

	response.write "</webActivity>"
	
%>	