<%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/9 上午 10:55:22
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\web\member\member_forget_act.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("submit.x", "cancel.x", "account")
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
%><!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<% 
sub StrRandomize(strSeed)
	dim i,nSeed 
	nSeed = CLng(0)
	for i = 1 To Len(strSeed)
		nSeed = nSeed Xor ((256 * ((i - 1) Mod 4) * AscB(Mid(strSeed, i, 1))))
	next
	Randomize nSeed
end sub

Function GeneratePassword(nLength)
	Dim i, bMadeConsonant, c, nRnd

	Const strDoubleConsonants = "npqrstuvwxy"
	Const strConsonants = "abcdefghijkm"
	Const strVocal = "3456789"

	GeneratePassword = ""
	bMadeConsonant = False

	For i = 0 To nLength
		nRnd = Rnd
		If GeneratePassword <> "" AND (bMadeConsonant <> True) AND (nRnd < 0.25) Then
			c = Mid(strDoubleConsonants, Int(Len(strDoubleConsonants) * Rnd + 1), 1)
			c = c & c
			i = i + 1
			bMadeConsonant = True
		Else
			If (bMadeConsonant <> True) And (nRnd < 0.75) Then
				c = Mid(strConsonants, Int(Len(strConsonants) * Rnd + 1), 1)
				bMadeConsonant = True
			Else
				c = Mid(strVocal,Int(Len(strVocal) * Rnd + 1), 1)
				bMadeConsonant = False
			End If
		End If
		GeneratePassword = GeneratePassword & c
	Next

	If Len(GeneratePassword) > nLength Then
		GeneratePassword = Left(GeneratePassword, nLength)
	End If
End Function
  '取得輸入
  submit = request("submit.x")
  cancel = request("cancel.x")
  
  '取消
  if trim(cancel) <> "" then
    response.write "<script language='JavaScript'>location.replace('default.asp');</script>"
  
  '確定
  'if trim(submit) <> "" then
  else
    '取得輸入
    account = request("account")
    realname = request("realname")
    email = request("email")
    
    '檢查輸入
    if trim(account) = "" then
      response.write "<script language='JavaScript'>alert('請輸入帳號 !!');history.go(-1);</script>"
    elseif trim(realname) = "" then
      response.write "<script language='JavaScript'>alert('請輸入姓名 !!');history.go(-1);</script>"
    elseif trim(email) = "" then
      response.write "<script language='JavaScript'>alert('請輸入電子信箱 !!');history.go(-1);</script>"
    '檢查email格式是否正確
	elseif len(email) <= 3 and InStr(email, "@") <= 0 and InStr(email, ".") <= 0 then
      '不正確...跳回
		response.write "<script language='JavaScript'>alert('您的電子信箱格式輸入錯誤 !!');history.go(-1);</script>"
	else
      '正確

    '檢查帳號是否存在
		set ts = conn.Execute("select count(*) from Member where account = '" & account & "'")
		if ts(0) < 1 then
			response.write "<script language='JavaScript'>alert('您的帳號或密碼錯誤 !!');history.go(-1);</script>"
		else
    
    '取得帳號基本資料
			sql = "select m.*, r.CtRootID from Member AS m LEFT JOIN CatTreeRoot AS r ON r.vGroup=m.mCode AND r.inUse='Y'" _
				& " where account = " & pkStr(account,"")
			set rs = conn.execute(sql)
    
    '檢查姓名
			if trim(rs("realname")) <> trim(realname) then
				response.write "<script language='JavaScript'>alert('您的姓名或電子信箱錯誤 !!');history.go(-1);</script>"
			elseif trim(rs("email")) <> trim(email) then
				response.write "<script language='JavaScript'>alert('您的姓名或電子信箱錯誤 !!');history.go(-1);</script>"
			else
	'產生密碼
				StrRandomize(right(CStr(Now),4) )
				passwd1 = GeneratePassword(6)

    '將新密碼存入資料庫
				sql = "update Member set passwd='" & passwd1 & "' where account='" & account & "'"
				conn.execute(sql)

	'寄出密碼
				Set myMail = CreateObject("CDO.Message")
				myMail.Subject = "財政部全球資訊網 - 會員密碼提醒"
				myMail.From = "財政部全球資訊網 <websys@mof.gov.tw>"
				myMail.To = email 
				myMail.HTMLBody = "請使用系統新產生的密碼登入，登入後可自行修改密碼。<p>" & realname & "，您的新密碼：<b>" & passwd1 & "<b>"  
				myMail.Send
				set myMail = nothing

    '導回首頁
				response.write "<script language='JavaScript'>alert('新密碼已寄出 !!');location.replace('default.asp');</script>"
			end if
		end if
	end if
  end if
%>