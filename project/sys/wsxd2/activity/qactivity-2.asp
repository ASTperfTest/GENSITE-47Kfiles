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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\activity\qactivity-2.asp"

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
		    	<li><a href="/sp.asp?xdURL=activity/qactivity-1.asp&mp=<%=mp%>">活動簡介</a></li>
		      <li class="now"><a href="/sp.asp?xdURL=activity/qactivity-2.asp&mp=<%=mp%>">活動辦法</a></li>
		      <!--li><a href="/sp.asp?xdURL=activity/qactivity-3.asp&mp=<%=mp%>">活動獎項</a></li-->
		      <li><a href="/sp.asp?xdURL=activity/qactivity-4.asp&mp=<%=mp%>">活動說明</a></li>
		      <li><a href="/sp.asp?xdURL=activity/qactivity-5.asp&mp=<%=mp%>">登入參加問卷調查</a></li>
		    </ul>
		
				<div class="content">		    
		    	<h4>活動辦法</h4>		
		      <div class="list">
		      	<ol>
		          <li>1.活動時間為：<strong class="red01">2008/6/2中午12點 ~ 2008/6/13中午12點</strong>（依實際活動上線時間為主）。</li>
		          <li>2.本次調查活動<strong class="red01">限會員參加，每個會員限參加一次</strong>。</li>
		          <li>3.活動結束後1週內，將公開由電腦隨機抽出25名得獎者，並於隔日在「農業知識入口網（http://kmweb.coa.gov.tw/）進行公告。</li>
		          <li>4.公佈名單後1週內，將以e-mail及電話方式通知得獎者，中獎民眾需於期限內依通知進行身分確認。各獎項將於公佈名單2週內陸續寄送。</li>
		          <li>5.中獎名單公佈後，得獎者需於指定時間內完成身份的確認，逾期或身份不符者皆視為放棄得獎資格，不另候補。</li>
		          <li>6.得獎者需正確填寫領獎收據，並附上身份證影本以便進行核對領獎身分。得獎者個人資料僅供本次核對領獎身份用，不予退還。</li>
		          <li>7.本次活動得獎贈品僅限郵寄台、澎、金、馬地區。</li>
		          <li>8.依稅法規定，贈品價值逾新台幣13,333元時，得獎者需負擔15 % 贈品稅，本活動贈品未超過此金額，得獎者無需負擔贈品稅。            </li>
		          <li>9.同一得獎者限得獎一次。若重複中獎，主辦單位得以取消其資格、採順位候補。</li>
		          <li>10.如有任何因電腦、網路、電話、技術或不可歸責於主辦單位之事由，而使參加者所寄出或登錄之資料有遲延、遺失、錯誤、無法辨識或毀損之情況，主辦單位不負任何法律責任，參加者亦不得因此異議。</li>
		          <li> 11.本活動規則及注意事項載明在本活動網頁中，參加者於參加本活動之同時，即同意接受本活動規則及注意事項之規範 ，如有違反本活動辦法及注意事項，主辦單位得取消其資格，並對於任何破壞本活動之行為保留法律追訴權。</li>
		          <li> 12.活動期間為維護比賽公平性，禁止網友從事外掛程式或任何影響活動公平之行為，一但查證告發，主辦單位有權不需說明，直接刪除其抽獎和得獎資格，並可採順位候補不另行通知。</li>
		          <li>13.本活動所涉及的個人資料蒐集、運用，以及關於個人資料的保護，將適用農委會網站的隱私權保護政策。</li>
		          <li>14.主辦單位保留贈品、活動期間的修改、及中獎資格審核最終權利。 </li>
		        </ol>
					</div>
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