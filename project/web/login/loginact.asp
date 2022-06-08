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
progPath="D:\hyweb\GENSITE\project\web\subject\loginact.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("account", "type2")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
genParamsMessage=now() & vbTab & "Error(1):REQUEST變數含特殊字元"

'-------- (可修改)只檢查單引號，如 Request 變數未來將要放入資料庫，請一定要設定(防 SQL Injection) --------
sqlInjParamsArray=array()
sqlInjParamsPattern=array("'")	'#### 要檢查的 pattern(程式會自動更新):DB ####
sqlInjParamsMessage=now() & vbTab & "Error(2):REQUEST變數含單引號"

'-------- (可修改)只檢查 HTML標籤符號，如 Request 變數未來將要輸出，請設定 (防 Cross Site Scripting)--------
xssParamsArray=array("O_URL")
xssParamsPattern=array("<", ">", "%3C", "%3E",";","%3B","""","%22")	'#### 要檢查的 pattern(程式會自動更新):HTML ####
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
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<?xml version="1.0"  encoding="utf-8" ?>
<html xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:hyweb="urn:gip-hyweb-com" xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>農業知識入口網 －小知識串成的大力量－/</title>
</head>
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/web.sqlInjection.inc" -->

<%
	Function nullText(xNode)
  		On Error Resume Next
  		xstr = ""
  		xstr = xNode.text
  		nullText = xStr
	End Function

	mp = getMPvalue() 
	
	account2 = stripHTML( request("account") )
	passwd2 = stripHTML( request("passwd") )
	O_URL = stripHTML( request("O_URL") )
	'type2 = stripHTML( request("type2") )
		
	Set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = True	
	Set oxsl = server.createObject("microsoft.XMLDOM")
	
	xv = oxml.load( session("myXDURL") & "/wsxd2/xdlogin.aspx?mp=" & mp & "&account2=" & account2 & "&passwd2=" & passwd2 )

  If oxml.parseError.reason <> "" Then 
   	Response.Write("htPageDom parseError on line " &  oxml.parseError.line)
   	Response.Write("<BR>Reason: " &  oxml.parseError.reason)
   	Response.End()
  End If
    
  login = nullText(oxml.selectSingleNode("//login/status"))
	'response.write login
	'response.end
  If login = "True" Then
	'2010-6-17::判斷會員 Email 是否通過認證
	mcode = nullText(oxml.selectSingleNode("//login/mcode"))
	If mcode <> "Y" then 
		updateMemID = stripHTML( nullText(oxml.selectSingleNode("//login/memID")) )
%>
   	<script language="javascript">
     		alert("會員 Email 尚未通過認證，請重新認證。");
			window.location.href = "/sp.asp?xdURL=Coamember/member_reVerifyEmail.aspx&mp=1&updateMemID=<%=updateMemID%>";
   	</script>   
<%
		Response.End()
	End If
  
   	Session("memID") 					= stripHTML( nullText(oxml.selectSingleNode("//login/memID")) )
   	Session("gstyle") 				= stripHTML( nullText(oxml.selectSingleNode("//login/gstyle")) )
   	Session("memName") 				= stripHTML( nullText(oxml.selectSingleNode("//login/memName")) )
   	Session("memNickName") 		= stripHTML( nullText(oxml.selectSingleNode("//login/memNickName")) )
   	Session("memLoginCount") 	= stripHTML( nullText(oxml.selectSingleNode("//login/memLoginCount") ) )       	   	
   	'response.write "login"  
		
		'---login 成功---
		
		sql = "SELECT LastLoginTime FROM ActivityMemberNew WHERE MemberId = '" & Session("memID") & "'"		
		Set loginrs = conn.execute(sql)
		
		If Not loginrs.Eof Then
			If IsNull(loginrs("LastLoginTime")) Then
				sql = "UPDATE ActivityMemberNew Set LoginGrade = LoginGrade + 2, LastLoginTime = GETDATE() WHERE MemberId = '" & Session("memID") & "'"
			Else						
				If DateDiff("d", loginrs("LastLoginTime"), Date) = 0 Then
					sql = "UPDATE ActivityMemberNew Set LastLoginTime = GETDATE() WHERE MemberId = '" & Session("memID") & "'"
				Else
					sql = "UPDATE ActivityMemberNew Set LoginGrade = LoginGrade + 2, LastLoginTime = GETDATE() WHERE MemberId = '" & Session("memID") & "'"
				End If			
			End If
		Else
			'---no record, insert a new record---
			sql = "INSERT INTO ActivityMemberNew (MemberId, LoginGrade, LastLoginTime) VALUES ('" & Session("memID") & "', 2, GETDATE()) "		
		End If		
		'response.write sql
		conn.execute(sql)		
		
		'---20080915---vincent---加入目前頁面的url, 登入後導到目前頁面---
%>
   	<script language="javascript">
     		alert("登入成功!!");
			var redirectUrl="<%=O_URL%>";
			redirectUrl=redirectUrl.replace("http://","");
			redirectUrl=redirectUrl.substr(redirectUrl.indexOf('/'));
			window.location.href = "/AspSession.asp?redirecturl="+redirectUrl;
   	</script>   
<%
  Else
   	Session("memID") = ""
   	Session("gstyle") = "" 
   	Session("memName") = ""	
   	Session("memNickName") = ""
   	Session("memLoginCount") = ""
   	
   	logintype = nullText(oxml.selectSingleNode("//login/type"))
   	If logintype = "1" Then
%>
   		<script language="javascript">
	   		alert("帳號或密碼錯誤!!");
   			//window.location="/mp.asp?mp=1";
   			history.back();
   		</script>
<%
		ElseIf logintype = "2" Then
%>
   		<script language="javascript">
	   		alert("此帳號已被停權，若有問題請聯繫系統管理員!!");
   			//window.location="/mp.asp?mp=1";
   			history.back();
   		</script>
<%			
		End If
  End If
  'response.redirect "mp.asp?mp=" & mp	
%>
