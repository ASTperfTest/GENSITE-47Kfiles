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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\Coamember\mailbody_S.inc"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("account")
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
%><!--#include virtual = "/inc/checkGIPAPconfig.inc" -->
<%


dim SMTPServer
dim SMTPServerPort
dim SMTPSsendusing

dim SMTPusername
dim SMTPpassword
dim SMTPauthenticate

'SMTPServer = "mail.coa.gov.tw"
'SMTPServerPort = 25
'SMTPSsendusing = 1
'SMTPSsendusing = 2

SMTPServer = getGIPApconfigText("EmailServerIp")
SMTPServerPort = getGIPApconfigText("EmailServerPort")
SMTPSsendusing = getGIPApconfigText("EmailServerSendUsing")

Function Send_Email_authenticate(S_Email,R_Email,Re_Sbj,Re_Body, SMTPusername, SMTPpassword, SMTPauthenticate)
    Set objEmail = CreateObject("CDO.Message")
    objEmail.bodypart.Charset = "UTF-8"'response.charset
    objEmail.From       = S_Email
    objEmail.To         = R_Email
    objEmail.Subject    = Re_Sbj
    objEmail.HTMLbody   = Re_Body
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = SMTPSsendusing
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPServerPort

    if SMTPusername <> "" then
        objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPusername
        objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPpassword
        objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = SMTPauthenticate
    end if

    objEmail.Configuration.Fields.Update
    objEmail.Send
    set objEmail=nothing
End Function

Function Send_Email(S_Email,R_Email,Re_Sbj,Re_Body)
    dim SMTPusername
    dim SMTPpassword
    dim SMTPauthenticate

    'SMTPusername = "taft_km"
    'SMTPpassword = "P@ssw0rd"
    'SMTPauthenticate = "1"
	
	SMTPusername = getGIPApconfigText("Emailsendusername")
	SMTPpassword = getGIPApconfigText("Emailsendpassword")
	SMTPauthenticate = getGIPApconfigText("Emailsmtpauthenticate")

    call Send_Email_authenticate(S_Email,R_Email,Re_Sbj,Re_Body, SMTPusername, SMTPpassword, SMTPauthenticate)
End Function


ePaperTitle = "農業知識入口網會員申請通知"

key = session("WebDESKey")	
Set DesCrypt = New Cls_DES
token = DesCrypt.DES(request("email"),key,0)	'加密
Set DesCrypt = Nothing

Body = "親愛的 " & Request("realname") & "您好:<br /><br />" & vbCrLf & vbCrLf
Body = Body & "「農業知識入口網」已收到您的申請資料。您的會員帳號密碼如下:<br />" & vbCrLf
Body = Body & "您的帳號: " & Request("account") & "<br />" & vbCrLf
Body = Body & "您的密碼: " & Request("passwd") & "<br /><br />" & vbCrLf & vbCrLf
Body = Body & "請您點選信中的網址以確認您的 E-mail ：<br /><br />" 
Body = Body & "<a href='" & session("WebURL") & "ConfirmEmail.asp?email=" & request("email") & "&token=" & token & "&id=" & request("account") & "'>確認Email</a><br /><br />" 
Body = Body & "請確實保管本信件，並記妥個人申請帳號密碼，在學者會員審核通過前，您仍可使用一般會員權益。<br />" & vbCrLf
Body = Body & "並感謝您申請「農業知識入口網」的學者會員。<br />" & vbCrLf
Body = Body & "為了維護會員的權益，我們將對於您提出的申請，做資料上的核對與確認。<br />" & vbCrLf
Body = Body & "通過審核後，我們將立刻通知您。<br />" & vbCrLf
Body = Body & "    謝謝!<br />" & vbCrLf
Body = Body & "                                      敬祝平安<br />" & vbCrLf
Body = Body & "                                                系統管理員 敬上<br />" & vbCrLf

'S_Email = "km@mail.coa.gov.tw"
S_Email = """"&getGIPApconfigText("EmailFromName")&""" <"&getGIPApconfigText("EmailFrom")&">"
R_Email = Request("email")
		
Call Send_Email(S_Email,R_Email,ePaperTitle,Body)

ePaperTitle = "農業知識入口網會員申請通知"

Body = ""
Body = "親愛的 管理者您好:" & vbCrLf & vbCrLf
Body = Body & "「農業知識入口網」有新的學者會員申請資料。會員帳號如下:" & vbCrLf
Body = Body & "帳號: " & Request("account") & vbCrLf
Body = Body & "    謝謝!" & vbCrLf & vbCrLf
Body = Body & "                                      敬祝平安" & vbCrLf
Body = Body & "                                                系統管理員 敬上" & vbCrLf

'S_Email = "km@mail.coa.gov.tw"
S_Email = """"&getGIPApconfigText("EmailFromName")&""" <"&getGIPApconfigText("EmailFrom")&">"
R_Email = Request("email")
		
Call Send_Email(S_Email,S_Email,ePaperTitle,Body)




'Dim Body               
'Set mail = Server.CreateObject("CDONTS.NewMail")
'mail.To = Request("email")
'mail.From = "km@mail.coa.gov.tw"
'mail.Subject = "農業知識入口網會員申請通知"
'Body = "親愛的 " & Request("realname") & "您好:" & vbCrLf & vbCrLf
'Body = Body & "「農業知識入口網」已通過您的申請資料。您的會員帳號密碼如下:" & vbCrLf
'Body = Body & "您的帳號: " & Request("account") & vbCrLf
'Body = Body & "您的密碼: " & Request("passwd") & vbCrLf & vbCrLf
'Body = Body & "請確實保管本信件，並記妥個人申請帳號密碼" & vbCrLf
'Body = Body & "有了您的參與將使得這個農業知識平台更理想更茁壯。歡迎您正式成為我們的會員，會員服務中心的網址是熱切期盼您提供寶貴的意見與建議。" & vbCrLf
'Body = Body & "    謝謝!" & vbCrLf & vbCrLf
'Body = Body & "     感謝您申請「農業知識入口網」的學者會員。" & vbCrLf
'Body = Body & "為了維護會員的權益，我們將對於您提出的申請，做資料上的核對與確認。" & vbCrLf
'Body = Body & "通過審核後，我們將立刻通知您。" & vbCrLf
'Body = Body & "    謝謝!" & vbCrLf
'Body = Body & "                                      敬祝平安" & vbCrLf
'Body = Body & "                                                系統管理員 敬上" & vbCrLf
'mail.Body = Body
'mail.Send
'Set mail = Nothing

'Set mail = Server.CreateObject("CDONTS.NewMail")
'mail.To = "km@mail.coa.gov.tw"
'mail.To = "candan@mail.hyweb.com.tw"
'mail.From = Request("email")
'mail.Subject = "農業知識入口網會員申請通知"
'Body = "親愛的 管理者您好:" & vbCrLf & vbCrLf

'Body = Body & "「農業知識入口網」有新的申請資料。會員帳號如下:" & vbCrLf
'Body = Body & "帳號: " & Request("account") & vbCrLf
'Body = Body & "    謝謝!" & vbCrLf & vbCrLf

'Body = Body & "                                      敬祝平安" & vbCrLf
'Body = Body & "                                                系統管理員 敬上" & vbCrLf
'mail.Body = Body
'mail.Send
'Set mail = Nothing
%>