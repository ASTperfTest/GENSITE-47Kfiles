<%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/10 上午 09:59:38
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\webActivity\actPage5.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("id", "mp")
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
<%
	Dim activityId : activityId = ""
	Dim flag : flag = true
	Dim memberId : memberId = ""
	Dim mp : mp = ""
	
	activityId = request.queryString("id")
	memberId = session("memID")
	mp = request.querystring("mp")
	
	if activityId = "" or isempty(activityId) then  flag = false
	if not isnumeric(activityId) then flag = false
	if isnumeric(activityId) And ( instr(activityId, ".") > 0 or instr(activityId, "-") > 0 ) then flag = false
	if not isnumeric(mp) then mp = "1"
	
	if not flag then
		response.write "<script>"
		response.write "alert('無相關問卷調查');"
		response.write "window.location.href='/';"
		response.write "</script>"
	else
		sql = "SELECT * FROM m011 WHERE GETDATE() BETWEEN m011_bdate AND m011_edate AND m011_online = '1' AND m011_subjectid = " & activityId
		set rs = conn.execute(sql)
		If rs.EOF Then 
			flag = false		
			response.write "<script>"
			response.write "alert('無相關問卷調查');"
			response.write "window.location.href='/';"
			response.write "</script>"
		end if
		rs.close
		set rs = nothing		
	end if
	Dim returnUrl
	returnUrl = "loginact.asp?redirecturl=" & server.urlencode("/sp.asp?xdURL=webActivity/QuestionPage.asp;mp="& mp & ";id=" & activityId)
	if flag then 		
		if memberId <> "" then			
			response.write "<script>window.location.href='/sp.asp?xdURL=webActivity/QuestionPage.asp&mp=" & mp & "&id=" & activityId & "'</script>"
			'response.end
		else
%>
			<div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="#">問卷調查</a></div>
		
			<div class="actblock">
				<h3>農業知識入口網－網站使用者需求暨滿意度調查－</h3>
			  <ul class="function">
					<!--li><a href="/sp.asp?xdURL=webActivity/actPage1.asp&mp=<%=mp%>&id=<%=activityId%>">活動簡介</a></li>
			    <li><a href="/sp.asp?xdURL=webActivity/actPage2.asp&mp=<%=mp%>&id=<%=activityId%>">活動辦法</a></li>
			    <li><a href="/sp.asp?xdURL=webActivity/actPage3.asp&mp=<%=mp%>&id=<%=activityId%>">活動獎項</a></li>
			    <li><a href="/sp.asp?xdURL=webActivity/actPage4.asp&mp=<%=mp%>&id=<%=activityId%>">活動說明</a></li>
			    <li class="now"><a href="/sp.asp?xdURL=webActivity/actPage5.asp&mp=<%=mp%>&id=<%=activityId%>">登入參加問卷調查</a></li-->
			  </ul>

				<div class="content">
					<h4>會員登入</h4>
					<p>本次活動限定會員參加，參加調查前，<strong class="red01">請您先登入或加入會員</strong>。，謝謝您的配合。</p>
					<hr/>
					<div class="ce">
						<a href="http://kmbeta.coa.gov.tw/sp.asp?xdURL=coamember/member_Join.asp&mp=1" target="_blank"><img src="/xslGip/style1/images2/activity_register.gif" border="0" /></a>
						<form class="a01" name="iForm" action="<%=returnUrl%>" method="post">
							<br/><br/><br/><br/><br/><br/><br/>
							<table border="0" cellpadding="0" cellspacing="0" summary="內容表格">
						  <tr>
								<th scope="row">會員帳號：</th>
								<td><input type="text" name="account2" size="40" value="請輸入您註冊的帳號" onclick="this.value=''" onfocus="this.value=''" /></td>
			        </tr>
						  <tr>
								<th scope="row">會員密碼：</th>
								<td><input type="password" name="passwd2" size="40" value="" onclick="this.value=''" onfocus="this.value=''" /></td>
						  </tr>
							</table>
							<br/>
							<input name="Submit" type="submit" class="button" value="送出" />
							<input name="Reset" type="reset" class="button" value="重設" />
							<br/><br/><br/><br/>
						</form>
					</div>
					<br/><br/><br/><br/>
				</div>
			</div>
<%		
		end if
	end if
%>