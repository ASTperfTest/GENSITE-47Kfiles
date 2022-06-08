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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\activity\qactivity-1.asp"

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
	ActivityId = Application("ActivityId")
	sql = "SELECT * FROM m011 WHERE GETDATE() BETWEEN m011_bdate AND m011_edate AND m011_subjectid = " & ActivityId
	set rs = conn.execute(sql)
	If rs.EOF Then
		response.write "<script>window.location.href='/'</script>"
	Else
		mp = request("mp")
		if not isnumeric(mp) then mp = "1"
%>

		<td class="center">			
			<div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="#">問卷調查</a></div>									
			
			<div class="actblock">	
				<h3>農業知識入口網－97年網站使用者需求(滿意度)問卷調查－</h3>
		    <ul class="function">
		    	<li class="now"><a href="/sp.asp?xdURL=activity/qactivity-1.asp&mp=<%=mp%>">活動簡介</a></li>
		      <li><a href="/sp.asp?xdURL=activity/qactivity-2.asp&mp=<%=mp%>">活動辦法</a></li>
		      <!--li><a href="/sp.asp?xdURL=activity/qactivity-3.asp&mp=<%=mp%>">活動獎項</a></li-->
		      <li><a href="/sp.asp?xdURL=activity/qactivity-4.asp&mp=<%=mp%>">活動說明</a></li>
		      <li><a href="/sp.asp?xdURL=activity/qactivity-5.asp&mp=<%=mp%>">登入參加問卷調查</a></li>
		    </ul>
		
				<div class="content">
					<h4>活動簡介</h4>		
					<script type="text/javascript">
						AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0','width','545','height','170','title','問卷調查','src','xslgip/style1/images2/banner_a','quality','high','pluginspage','http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash','movie','xslgip/style1/images2/banner_a' ); //end AC code
					</script>
					<noscript>
						<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="545" height="170" title="問卷調查">
		         	<param name="movie" value="xslgip/style1/images2/banner_a.swf" />
		          <param name="quality" value="high" />
		          <embed src="xslgip/style1/images2/banner_a.swf" quality="high" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="545" height="170"></embed>
						</object>
					</noscript>
					<br/><br/>
					<p>親愛的農業知識入口網的使用者，您好：</p>
					<p>農業知識入口網在去年進行了網站的全面改版。為了持續提供您更豐富的內容，我們再次推出「97年網站使用者需求(滿意度)問卷調查」調查活動。</p>
					<p>自<strong class="red01">2008/6/2起-2008/6/13</strong>止，只要您登入會員，並填寫完成「調查問卷」，就有機會抽中攜帶方便的<strong class="red01">8G隨身碟2名</strong>和<strong class="red01">1G迷你隨身碟23名</strong>，<strong class="red01">共有25名</strong>得獎機會，這樣好康的機會不可多得喔!!</p>
					<p>本活動參加對象需為知識入口網之會員，得獎名單將於活動結束1週內公佈於農業知識入口網，歡迎民眾一起來共襄盛舉。</p>
					<img src="xslgip/style1/images2/pactivity_prize.gif" />
					<table border="0" cellpadding="0" cellspacing="0" class="acttable" summary="內容表格">
					<tr>
						<th scope="col"> 獎項 </th>
						<th scope="col">獎項</th>
						<th scope="col">市價</th>
						<th scope="col">名額</th>
					</tr>
					<tr>
						<td> 特獎 </td>
						<td>8G隨身碟</td>
						<td>1,500</td>
						<td>2</td>
					</tr>
					<tr>
		  			<td>普獎</td>
		  			<td>1G隨身碟</td>
		  			<td>500</td>
		  			<td>23</td>
					</tr>
					</table>
					<div class="space"></div>
					<br/><br/>
					<center><a href="/sp.asp?xdURL=activity/qactivity-5.asp&mp=<%=mp%>"><img src="xslgip/style1/images2/activity_go.gif" border="0" /></a></center>
				</div>				
			</div>			
		</td>
<%
	End If
	set rs = nothing
%>