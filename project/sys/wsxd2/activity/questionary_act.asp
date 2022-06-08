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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\activity\questionary_act.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("mp")
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
%><!--#Include virtual = "/inc/client.inc" -->
<%
	Dim flag
	flag = true
	
	ActivityId = Application("ActivityId")
	mp = request("mp")
	if not isnumeric(mp) then mp = "1"

	'---check the activity is in the activity time---
	sql = "SELECT * FROM m011 WHERE GETDATE() BETWEEN m011_bdate AND m011_edate AND m011_subjectid = " & ActivityId
	set rs = conn.execute(sql)
	If rs.EOF Then
		flag = false
		response.write "<script>window.location.href='/'</script>"		
	End If
	Set rs = nothing
	
	Dim memberId
	'---check session---
	If flag Then
		If IsNull(Session("memID")) Or Session("memID") = "" Then
			flag = false
			response.write "<script>"
			response.write "alert('session已失效，請重新登入後再次填寫問卷！');"
			response.write "window.location.href='/sp.asp?xdURL=activity/qactivity-5.asp&mp=" & mp & "';</script>"			
		Else
			memberId = Session("memID")
		End If		
	End If
	
	'---check the user have filled this question---
	If flag Then
		sql = "SELECT * FROM m014 WHERE m014_subjectid = " & ActivityId & " AND m014_name = '" & memberId & "'"
		Set memrs = conn.execute(sql)
		If Not memrs.Eof Then
			flag = false
			response.write "<script>alert('你已做過此調查！');history.go(-1);</script>"
			'if onlyonce = "1" and email <> "" then
			'sql = "select m014_email from m014 where m014_subjectid = " & subjectid & " and m014_email = '" & email & "'"
			'set rs = conn.execute(sql)
			'if not rs.eof then
			'	response.write "<script language='javascript'>alert('你已做過此調查！');history.go(-1);</script>"
			'	response.end
			'end if
			'end if
		End If
	End If
	Set memrs = nothing
	
	'---check the all the answer have been filled---
	If flag Then				
		sql2 = "SELECT m012_questionid FROM m012 WHERE m012_subjectid = " & ActivityId & " ORDER BY 1"		
		Set rs2 = conn.execute(sql2)
		While Not rs2.Eof
			questionid = rs2(0)
			If request("answer" & questionid) = "" Then
				flag = false				
				response.write "<script language='javascript'>alert('請選擇第 " & questionid & " 題答案！');history.go(-1);</script>"								
			End If
			rs2.MoveNext
		Wend	
		Set rs2 = nothing
	End If
	
	'---選出開放式答題的項目---m015---
	If flag Then				
		sql = "SELECT m015_questionid, m015_answerid FROM m015 WHERE m015_subjectid = " & ActivityId
		Set rs3 = conn.execute(sql)
		While Not rs3.Eof
			open_str = open_str & "*" & rs3(0) & "*" & rs3(1) & "*,"
			rs3.MoveNext
		Wend
		Set rs3 = nothing
	End If
	
	'---記錄填寫過此問卷的ID---m014---
	If flag then					
		sql = ""
		If memberId <> "" Then
			set rs3 = conn.execute("select isNull(max(m014_id), 0) + 1 from m014")
			m014_id = rs3(0)
			sql = "" & _
						" insert into m014 ( " & _
						" m014_id, " & _
						" m014_name, " & _
						" m014_sex, " & _
						" m014_email, " & _
						" m014_age, " & _
						" m014_addrarea, " & _
						" m014_familymember, " & _
						" m014_money, " & _
						" m014_job, " & _
						" m014_edu, " & _
						" m014_pflag, " & _
						" m014_reply, " & _
						" m014_subjectid, " & _
						" m014_polldate " & _
						" ) values ( " & _
						m014_id & ", " & _
						" '" & memberId & "', " & _
						" '', " & _
						" '', " & _
						" '', " & _
						" '', " & _
						" '', " & _
						" '', " & _
						" '', " & _
						" '', " & _
						" '0', " & _
						" '', " & _
						ActivityId & ", " & _
						" getdate() " & _
						" ); "
		Else
			m014_id = 0
		End If
	End If
	
	If flag Then
		'---save the answers into db---
		sql2 = "SELECT m012_questionid, m012_textarea FROM m012 WHERE m012_subjectid = " & ActivityId & " ORDER BY 1"
		Set rs = conn.execute(sql2)
		While Not rs.Eof
			
			questionid = rs(0)
			textareaid = rs(1)
			answer_a = split(request("answer" & questionid), ",")
			
			For i = 0 to ubound(answer_a)
			
				'---update the counter of question---add by 1---
				sql = sql & _
							" UPDATE m013 set m013_no = m013_no + 1 WHERE " & _
							" m013_subjectid = " & ActivityId & " AND " & _
							" m013_questionid = " & questionid & " AND " & _
							" m013_answerid = " & answer_a(i) & "; "
				
				'---insert each answer into m018---
				sql = sql & _
						" INSERT INTO m018 ( " & _
						" m018_subjectid, " & _
						" m018_questionid, " & _
						" m018_answerid, " & _
						" m018_userid, " & _
						" m018_updatetime " & _
						" ) VALUES ( " & _
						ActivityId & ", " & _
						questionid & ", " & _								
						Trim(answer_a(i))& ", " & _
						m014_id & ", " & _
						" getdate() " & _
						" ); "			
									
				'---insert the 開放式答題的答案---m016---	
				If InStr(open_str, "*" & questionid & "*" & Trim(answer_a(i)) & "*") > 0 And _
					 Trim( request( "open_content" & questionid & "_" & Trim(answer_a(i)) ) ) <> "" Then					 						 	
					 	
					sql = sql & _
							" INSERT INTO m016 ( " & _
							" m016_subjectid, " & _
							" m016_questionid, " & _
							" m016_answerid, " & _
							" m016_userid, " & _
							" m016_content, " & _
							" m016_updatetime " & _
							" ) VALUES ( " & _
							ActivityId & ", " & _
							questionid & ", " & _
							Trim(answer_a(i))& ", " & _
							m014_id & ", " & _
							" '" & replace(trim( request( "open_content" & questionid & "_" & Trim(answer_a(i)) ) ), "'", "''") & "', " & _
							" getdate() " & _
							" ); "
				End If																				
			Next
			
			'---insert the textarea into db---
			If textareaid = "1" Then
			
				sql = sql & _
						"INSERT INTO m017 ( " & _
						" m017_subjectid, " & _
						" m017_questionid, " & _
						" m017_userid, " & _
						" m017_content, " & _
						" m017_updatetime " & _
						" ) VALUES ( " & _
						ActivityId & ", " & _
						questionid & ", " & _
						m014_id & ", " & _
						" '" & replace(trim(request("textarea" & questionid)), "'", "''") & "', " & _
						" getdate() " & _
							" ); "
			End If
			
			rs.movenext
		Wend			
		'response.write open_str
		'response.write sql
		Set rs = nothing
		'response.end
		conn.execute(sql)

		response.write "<script>alert('您的答題資料已經送出！');window.location.href('/sp.asp?xdURL=activity/qactivity-1.asp&mp=" & mp & "');</script>"
		'response.end			
		
	End If
	
	
%>