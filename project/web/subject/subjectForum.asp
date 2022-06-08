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
progPath="D:\hyweb\GENSITE\project\web\subject\subjectForum.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("starvalue", "icuitem")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "-", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
genParamsMessage=now() & vbTab & "Error(1):REQUEST變數含特殊字元"

'-------- (可修改)只檢查單引號，如 Request 變數未來將要放入資料庫，請一定要設定(防 SQL Injection) --------
sqlInjParamsArray=array("content")
sqlInjParamsPattern=array("'")	'#### 要檢查的 pattern(程式會自動更新):DB ####
sqlInjParamsMessage=now() & vbTab & "Error(2):REQUEST變數含單引號"

'-------- (可修改)只檢查 HTML標籤符號，如 Request 變數未來將要輸出，請設定 (防 Cross Site Scripting)--------
xssParamsArray=array("content")
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
	Response.Charset = "utf-8"
%>
<!--#Include virtual = "/inc/client.inc" -->
<%'xItem=83992&ctNode=1645&mp=86&kpi=0
starvalue=request("starvalue")
if starvalue="" then starvalue = 0 end if
content=request("content")
icuitem=request("icuitem")
'response.write starvalue&"<br>"
'response.write content&"<br>"
'response.write icuitem&"<br>"
'response.write Request.ServerVariables("HTTP_REFERER")&"<br>"

'檢查此文章在subjectForum 有沒有資料
sql_check = "SELECT * FROM SubjectForum where gicuitem = '"& icuitem &"'"
set rs_check = conn.execute(sql_check)
if rs_check.eof then
	sql_insert = "INSERT INTO SubjectForum (gicuitem) VALUES('"& icuitem &"')"
	conn.execute(sql_insert)
end if
'檢查該會員有無評論過
sql_check = "SELECT CuDTGeneric.iCUItem FROM CuDTGeneric INNER JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
sql_check = sql_check & "WHERE (CuDTGeneric.iEditor = '"& Session("memID") &"') AND (SubjectForum.ParentIcuitem = '"& icuitem &"') AND (SubjectForum.Status = 'Y')"
set rs_check2 = conn.execute(sql_check)
if not rs_check2.eof then
%>
	<script language="javascript">
		alert("您已評價過此篇文章，感謝您對主題館的支持!!");
		window.location.href ="<%=Request.ServerVariables("HTTP_REFERER")%>";
	</script> 
<%
	response.end
end if
if content = "" and starvalue = 0 then
%>
	<script language="javascript">
		alert("請發表你的意見與評價！");
		window.location.href ="<%=Request.ServerVariables("HTTP_REFERER")%>";
	</script> 
<%
	response.end
end if
'存意見內容
if content <> "" then
	sqlString = "INSERT INTO CuDTGeneric(iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, xBody, siteId) "
    sqlString = sqlString & "VALUES ('46', '2752', 'Y', '主題館意見-"& icuitem &"', '"& Session("memID") &"', '0', '"& content &"', '3')"
	sqlString = "set nocount on;"&sqlString&"; select @@IDENTITY as NewID"
	set rs_max = conn.execute(sqlString) '主表最新icuitem
	newicuitem = rs_max(0)
	sql_update = "UPDATE SubjectForum SET CommandCount=CommandCount+1 WHERE gicuitem='"& icuitem &"'"
	conn.execute(sql_update)
	sql_insert2 = "INSERT INTO SubjectForum(gicuitem, ParentIcuitem) VALUES ('"& newicuitem &"', '"& icuitem &"')"
	conn.execute(sql_insert2)'附表
end if	
'存評價分數
	if starvalue<>0 then
		sqlString = "INSERT INTO CuDTGeneric(iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, siteId) "
		sqlString = sqlString & "VALUES ('46', '2751', 'Y', '主題館評價-"& icuitem &"', '"& Session("memID") &"', '0', '3')"
		sqlString = "set nocount on;"&sqlString&"; select @@IDENTITY as NewID"
		set rs_max = conn.execute(sqlString) '主表最新icuitem
		newicuitem = rs_max(0)
		sql_update = "UPDATE SubjectForum SET GradeCount=GradeCount+"& CInt(starvalue) &",GradePersonCount = GradePersonCount + 1 WHERE gicuitem='"& icuitem &"'"
		conn.execute(sql_update)
		sql_insert2 = "INSERT INTO SubjectForum(gicuitem, GradeCount, GradePersonCount, ParentIcuitem) VALUES ('"& newicuitem &"','"&starvalue&"', '1','"& icuitem &"')"
		conn.execute(sql_insert2)'附表
        response.redirect "/Kpi/KpiInterShare.aspx?memberId=" & session("memID") & "&xItem=" & icuitem & "&icuitem=" & newicuitem & "&type=2" & "&refurl=" & server.urlencode(Request.ServerVariables("HTTP_REFERER"))
	end if

%>
	<script language="javascript">
		alert("評價成功，感謝您對主題館的支持!!");
		window.location.href ="<%=Request.ServerVariables("HTTP_REFERER")%>";
	</script> 