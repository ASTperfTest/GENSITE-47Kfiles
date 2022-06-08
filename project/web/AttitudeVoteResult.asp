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
progPath="D:\hyweb\GENSITE\project\web\AttitudeVoteResult.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("xItem", "ctNode", "mp", "kpi", "voteResult")
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
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<?xml version="1.0"  encoding="utf-8" ?>
<html xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:hyweb="urn:gip-hyweb-com" xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>農業知識入口網 －小知識串成的大力量－</title>
</head>
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/client.inc" -->
<!--#include file = "inc/web.sqlInjection.inc" -->

<%
	Dim ibaseDSD : ibaseDSD = 7
	Dim iCtUnit : iCtUnit = 2180

	'---此頁面不用利用網址帶參數, 若有帶, 則認為是hack---
	Dim hackFlag : hackFlag = false
	for each item in request.querystring
		hackFlag = true
	next
	if hackFlag then RedirectToPage "/"
			
	Dim xItem : xItem = request("xItem")
	Dim ctNode : ctNode = request("ctNode")
	Dim mp : mp = request("mp")
	Dim kpi : kpi = request("kpi")
	Dim voteResult : voteResult = request("voteResult")
	Dim voteDes : voteDes = request("voteDes")
	Dim memberId : memberId = session("memId")
	Dim sTitle : sTitle = ""
	Dim existFlag : existFlag = true
	
	if IsEmpty(memberId) then GoAlertMessage "請重新登入會員"
	if not IsNumeric(xItem) then RedirectToPage "/"
	if not IsNumeric(ctNode) then RedirectToPage "/"
	if not IsNumeric(mp) then RedirectToPage "/"
	if not IsNumeric(kpi) then RedirectToPage "/"
	if voteResult <> "A" and voteResult <> "B" and voteResult <> "C" and voteResult <> "D" and voteResult <> "E" then RedirectToPage "/"
	if Instr(voteDes, "%3c") > 0 then RedirectToPage "/"
	if Instr(voteDes, "%3d") > 0 then RedirectToPage "/"
	if Instr(voteDes, "%3e") > 0 then RedirectToPage "/"
			
	voteDes = stripHTML( voteDes )
	
	sql = "SELECT sTitle FROM CuDTGeneric WHERE iCUItem = " & xItem
	set rs = conn.execute(sql)
	if not rs.eof then 
		sTitle = rs("sTitle")
	else
		existFlag = false
	end if
	rs.close
	set rs = nothing
	if not existFlag then GoAlertMessage "文章不存在,無法投票"
	
	'---檢查是否自己投過票---
	Dim voteFlag : voteFlag = false
	sql = "SELECT iCUItem FROM CuDTGeneric WHERE iCtUnit = " & iCtUnit & " AND refID = " & xItem & " AND iEditor = '" & memberId & "' "	
	set rs = conn.execute(sql)
	if not rs.eof then
		voteFlag = true
	end if
	rs.close
	set rs = nothing
	if voteFlag then
		response.write "<script>alert('您已經投過票了');history.back();</script>"
		response.end
	end if
	sql = "set nocount on "
	sql = sql + " INSERT INTO CuDTGeneric(iBaseDsD, iCtUnit, fCTUPublic, sTitle, iEditor, iDept, topCat, refID, siteId, xBody) " & _
				"VALUES(" & iBaseDSD & ", " & iCtUnit & ", 'Y', '態度投票-" & sTitle & "', '" & memberId & "', '0', '" & voteResult & "', " & _
				xItem & ", '1', " & pkStr(voteDes, "") & ")"	
    sql = sql + " select scope_identity()  as idd "
	'response.write sql
	'response.end
	set rs =conn.Execute(sql)
	
	'---導到KPI記點---
	response.redirect "/Kpi/KpiInterShare.aspx?memberId=" & session("memID") & "&xItem=" & xItem & "&ctNode=" & ctNode & "&mp=" & mp & "&kpi=" & kpi & "&icuitem=" & rs("idd") & "&type=1" 
    rs.close
	set rs = nothing
  sub GoAlertMessage( str )
		response.write "<script language=""javascript"">alert('" & str & "');history.back();</script>"
		response.end
	end sub
	
	sub RedirectToPage( pagestr )
		response.redirect pagestr
		response.end
	end sub
	
	function pkStr (s, endchar)
		if s = "" then
			pkStr = "''" & endchar		
		else
			pos = InStr(s, "'")
			While pos > 0
				s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
				pos = InStr(pos + 2, s, "'")
			Wend
			pkStr = "N'" & s & "'" & endchar			
		end if
	end function
%>
