<%@ CodePage = 65001 %><%
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
progPath="D:\hyweb\GENSITE\project\web\epaper_act.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("textfield2", "CtRootID", "submit1.x", "submit2.x")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "-", "*", "/", "%",  "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
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
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#include file = "inc/web.sqlInjection.inc" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HEAD>
<TITLE>title</TITLE>
<META http-equiv="Content-Type" content="text/xml; charset=utf-8">
</HEAD>
<% On Error resume Next %>
<%

	'取得輸入
	email = pkStrWithSriptHTML(trim(request("textfield2")), "")
	CtRootID = pkStrWithSriptHTML( request.querystring("CtRootID"), "")
	submit1 = request.Form("submit1.x")
	submit2 = request.Form("submit2.x")
	
	'檢查是否輸入email
	if trim(email) = "null" then
	  response.write "<script language='JavaScript'>alert('請輸入您的 e-mail !!');history.go(-1);</script>"
	  response.end
	end if
	
	'訂閱電子報
	if trim(submit1) <> "" then
	  
	  '檢查是否已經訂閱
	  set ts = conn.Execute("select count(*) from Epaper where email = " & email & " and CtRootID = " & CtRootID )
		If Err.Number <> 0 or conn.Errors.Count <> 0 Then
			response.redirect "error.htm"
		end if		  
	  if ts(0) > 0 then
	    response.write "<script language='JavaScript'>alert('您的 e-mail 已經訂閱電子報 !!');history.go(-1);</script>"
	    response.end
	  end if
	  
    '檢查email格式是否正確
	  if len(email) > 3 and InStr(email, "@") > 0 and InStr(email, ".") > 0 then
	  
	    '正確...存入資料庫
	    sql = "insert into Epaper ( email, createtime, CtRootID) values ("& email & ", getdate()," & CtRootID & ")"
	    conn.execute(sql)
			
			If Err.Number <> 0 or conn.Errors.Count <> 0 Then
				response.redirect "error.htm"
			end if			
	    
	    response.write "<script language='JavaScript'>alert('恭喜您訂閱電子報成功!!');location.replace('mp.asp');</script>"
	    response.end
	    	    
	  else
	  
	    '不正確...跳回
	    response.write "<script language='JavaScript'>alert('您的 e-mail 格式輸入錯誤 !!');history.go(-1);</script>"
	    response.end
	    
	  end if

	end if
	
	'取消訂閱電子報
	if trim(submit2) <> "" then
	
	  '檢查是否已經訂閱
	  set ts = conn.Execute("select count(*) from Epaper where email = " & email & " and CtRootID = " & CtRootID)
		If Err.Number <> 0 or conn.Errors.Count <> 0 Then
			response.redirect "error.htm"
		end if		  
	  if ts(0) < 1 then
	    response.write "<script language='JavaScript'>alert('您的 e-mail 未曾訂閱電子報 !!');history.go(-1);</script>"
	    response.end
	  end if
	  
	  '刪除該筆email
	  conn.execute("delete from Epaper where email = " & email & " and CtRootID = " & CtRootID)
		If Err.Number <> 0 or conn.Errors.Count <> 0 Then
			response.redirect "error.htm"
		end if		  
	  response.write "<script language='JavaScript'>alert('已取消訂閱電子報 !!');location.replace('mp.asp');</script>"
	  response.end
	  
	end if
	
%>
