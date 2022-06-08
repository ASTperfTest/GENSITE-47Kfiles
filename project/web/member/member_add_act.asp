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
progPath="D:\hyweb\GENSITE\project\web\member\member_add_act.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("submit.x", "cancel.x", "htx_account", "htx_homeaddr", "htx_mobile", "htx_mcode")
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

  '取得輸入
  submit = request("submit.x")
  cancel = request("cancel.x")
  
  
  '取消
  if trim(cancel) <> "" then
    response.write "<script language='JavaScript'>location.replace('default.asp');</script>"
  
  '確定
  'elseif trim(submit) <> "" then
  else
    '取得輸入
    account = request("htx_account")
    passwd1 = request("htx_passwd1")
    passwd2 = request("htx_passwd2")
    realname = request("htx_realname")
    homeaddr = request("htx_homeaddr")
    phone = request("htx_phone")
    mobile = request("htx_mobile")
    email = request("htx_email")
    mcode = request("htx_mcode")

    'if bLen(account) > 30 then
     'response.write "<script language='JavaScript'>alert('帳號太長，限30個英文字母 !!');history.go(-1);</script>"
    '檢查輸入
    if trim(account) = "" then
      response.write "<script language='JavaScript'>alert('請輸入帳號 !!');history.go(-1);</script>"
    elseif trim(passwd1) = "" then
      response.write "<script language='JavaScript'>alert('請輸入密碼 !!');history.go(-1);</script>"
    elseif trim(passwd2) = "" then
      response.write "<script language='JavaScript'>alert('請輸入密碼確認 !!');history.go(-1);</script>"
    elseif trim(realname) = "" then
      response.write "<script language='JavaScript'>alert('請輸入姓名 !!');history.go(-1);</script>"
    elseif trim(email) = "" then
      response.write "<script language='JavaScript'>alert('請輸入電子信箱 !!');history.go(-1);</script>"
    elseif trim(mcode) = "" then
      response.write "<script language='JavaScript'>alert('請輸入身分群組 !!');history.go(-1);</script>"
    '檢查密碼
    elseif trim(passwd1) <> trim(passwd2) then
      response.write "<script language='JavaScript'>alert('兩次密碼不相同 !!');history.go(-1);</script>"
	else 
    '檢查是否已經有相同帳號
		set ts = conn.Execute("select count(*) from Member where account = '" & account & "'")
		if ts(0) > 0 then
			response.write "<script language='JavaScript'>alert('已經有相同帳號 !!');history.go(-1);</script>"
		else
	'檢查email格式是否正確
			if len(email) <= 3 and InStr(email, "@") <= 0 and InStr(email, ".") <= 0 then
      '不正確...跳回
				response.write "<script language='JavaScript'>alert('您的電子信箱格式輸入錯誤 !!');history.go(-1);</script>"
			else
      '正確
    
    '存入資料
			sql = "insert into Member (account,passwd,realname,homeaddr,phone,mobile,email,mcode,createtime,modifytime) values ('" & account & "','" & passwd1 & "','" & realname & "','" & homeaddr & "','" & phone & "','" & mobile & "','" & email & "','" & mcode & "',getdate(),getdate())"
			conn.execute(sql)
    
    '導回首頁
			response.write "<script language='JavaScript'>alert('恭喜您加入會員成功 !!');location.replace('default.asp');</script>"
			end if
		end if
	end if
  end if
%>