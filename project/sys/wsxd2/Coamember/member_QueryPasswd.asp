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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\Coamember\member_QueryPasswd.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array()
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
%>
<% response.expires = 0 %>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<%
	Server.ScriptTimeOut = 2000
	on error resume next

	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	
	Response.Buffer = true
	Dim Message
	
	email = trim(request("email"))
	email = replace(email,"'","''")
	realname = trim(request("realname"))
	realname = replace(realname,"'","''")

dim SMTPServer
dim SMTPServerPort
dim SMTPSsendusing

dim SMTPusername
dim SMTPpassword
dim SMTPauthenticate

SMTPServer = "mail.coa.gov.tw"
SMTPServerPort = 25
'SMTPSsendusing = 1
SMTPSsendusing = 2

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

    SMTPusername = "coakm"
    SMTPpassword = "1Q2w3e4r5T"
    SMTPauthenticate = "1"

    call Send_Email_authenticate(S_Email,R_Email,Re_Sbj,Re_Body, SMTPusername, SMTPpassword, SMTPauthenticate)
End Function
	
	
	'if realname = "" then
        '    Response.Write "<script language='javascript'>"
	'    Response.Write "alert('請輸入您的姓名!');"
	'    Response.Write "history.back();"
	'    Response.Write "</script>"
	'end if
	
	'if email = "" then
	'	Response.Write "<script language='javascript'>"
	'	Response.Write "alert('請輸入您的電子郵件信箱!');"
	'	Response.Write "history.back();"
	'	Response.Write "</script>"
	'	Response.end		
	'end if
	
	sql ="select * from member Where (realname='" & realname & "' or realname='" & Chg_UNI(realname) & "') and email='" & email & "'"
	'response.Write sql
	'response.End
    Set rs   = conn.Execute(sql)

    If not rs.eof Then   
    	response.Write "<HR>寄出通知信<HR>"
    	response.Write rs("passwd")
		'response.End
		
		'寄出通知信
		
		ePaperTitle = "農業知識入口網會員通知!"
		
		xmBody = "親愛的 " & trim(rs("realname")) & ":" & vbCrLf
		xmBody = xmBody & "您於" & date() & "申請查詢帳號密碼資料如下" & vbCrLf
		xmBody = xmBody & "帳號：" & rs("account") & vbCrLf
		xmBody = xmBody & "密碼：" & rs("passwd") & vbCrLf
		xmBody = xmBody & "請確實保管本信件，並記妥個人申請帳號密碼" & vbCrLf
		xmBody = xmBody & "謝謝！" & vbCrLf
		
		S_Email = """農委會農業知識入口網"" <km@mail.coa.gov.tw>"
		R_Email = email
		
		Call Send_Email(S_Email,R_Email,ePaperTitle,xmBody)
		'Dim Body
		'Set mail = Server.CreateObject( "CDONTS.NewMail" )                      
		'mail.To = email
		'mail.From = "taft_km@coa.gov.tw"
		'mail.Subject = "農業知識入口網會員通知!"
		'Body = "親愛的 " & trim(rs("realname")) & ":" & vbCrLf
		'Body = Body & "您於" & date() & "申請查詢帳號密碼資料如下" & vbCrLf
		'Body = Body & "帳號：" & rs("account") & vbCrLf
		'Body = Body & "密碼：" & rs("passwd") & vbCrLf
		'Body = Body & "請確實保管本信件，並記妥個人申請帳號密碼" & vbCrLf
		'Body = Body & "謝謝！" & vbCrLf
		'mail.Body = Body
		'mail.Send                                
		
		Response.Write "<html><body bgcolor='#ffffff'>"
  		Response.Write "<script language='javascript'>"

  		Response.Write "location.replace('sp.asp?xdURL=Coamember/member_GetpwOK.asp&mp=1');"
  		Response.Write "</script>"
  		Response.Write "</body></html>" 
	Else
		Response.Write "<html><body bgcolor='#ffffff'>"
		Response.Write "<script language='javascript'>"
		Response.Write "alert('您填寫的姓名或Email不正確，請重新輸入!');"
		Response.Write "history.back();"
		Response.Write "</script>"
		Response.Write "</body></html>"   
	End if
	
Function Chg_UNI(str)        'ASCII轉Unicode
	dim old,new_w,iStr
	old = str
	new_w = ""
	for iStr = 1 to len(str)
		if ascw(mid(old,iStr,1)) < 0 then
			new_w = new_w & "&#" & ascw(mid(old,iStr,1))+65536 & ";"
		elseif        ascw(mid(old,iStr,1))>0 and ascw(mid(old,iStr,1))<127 then
			new_w = new_w & mid(old,iStr,1)
		else
			new_w = new_w & "&#" & ascw(mid(old,iStr,1)) & ";"
		end if
	next
	Chg_UNI=new_w
End Function
%>
