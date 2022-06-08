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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\webActivity\actPage2.asp"

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
	Dim mp : mp = ""
	
	activityId = request.queryString("id")
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
		sql = "SELECT * FROM m011 WHERE GETDATE() BETWEEN m011_bdate AND m011_edate AND m011_online = '1' AND m011_subjectid = " & ActivityId
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
	if flag then 		
%>
<style type="text/css">
.activity1 {
	font-size: 12pt;
	color: #333;
	background-color: #FFF;
	font-family:Arial, Helvetica, sans-serif;
}
img{
	border:none;
}
h1{
	color:#F60;
	font-size:16px;
}
td{

}
th{

}
.content_table p{
	line-height:1.5em;
	margin:0 0 15px 0;
	font-size: medium;
}
.headerss{
	background:url(../images/header_bg.gif) repeat-x #fff ;
	height:165px;
	padding:0;
}
.content_table{
	margin:-3px 0 0 0 ;
}
.content{
	background:#ddf1f2;
	padding:15px;
}
.content_l{
	background:url(../images/content_bg_l.gif) left repeat-y #ddf1f2;
}
.content_r{
	background:url(../images/content_bg_r.gif) right repeat-y #ddf1f2;
}
#bt {
	text-align:center;
}
.tabs_ul {
	width: 100%;
	border-bottom: solid 1px #999;
	margin:0 0 30px 0;
	height:25px;
	padding:12px 0 0 0 ;
}
.tabs_ul  .tabs_ul_li {
	color: #333;
	float: left;
	height: 24px;
	line-height: 24px;
	overflow: hidden;
	position: relative;
	margin: 0 1px -1px 0;
	border: solid 1px #999;
	border-left: solid 1px #999;
	background: #e1e1e1;
	display: block;
	padding: 0 8px;
	
}
.tabs_ul  .tabs_ul_li_active  {
	color:#F60;
	float: left;
	height: 24px;
	line-height: 24px;
	overflow: hidden;
	position: relative;
	margin: 0 1px -1px 0;
	border: solid 1px #999;
	display: block;
	padding: 0 8px;
	background: #fff;
	border-bottom: solid 1px #fff;
	font-weight:bold;
}
.tabs_ul_li_active a {
	color:#F90;
	text-decoration:none;
	font-size: medium;
}
.tabs_ul_li a {
	color: #666;
	text-decoration:none;
	font-size: medium;
}

.tabs_ul_li:hover {
	color:#fff;
	background:#FC0;
}
.tabs_ul_li a:hover {
	color:#fff;
	font-size: medium;
}
#present{
	width:480px;
	margin:0 auto;
	padding:0 0 0 60px;
}
#present_list{
	background-image: url(../images/present_th_bg.gif);
	background-repeat: no-repeat;
	background-position: top;
	margin:0 auto;
}
#present_list th{
	color:#dcf1f2;
	height:24px;
	vertical-align:middle;
}
#present_list td{
	height:24px;
	vertical-align:middle;
	border-bottom:1px dashed #999;
}
#present ul li{
	font-size:10pt;
}
.footer2{
	background:url(../images/footer_bg.gif) repeat-x;
}
.link{
	border: #F30 dashed 1px;
	color:#F30;
	padding:2px;
	background:#fff;
	text-decoration:none;
}
.link:hover{
	color:#fff;
	background:#F30;
	border:#fff dashed 1px;
}
</style>
<div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="#">問卷調查</a></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="activity1">
  <tr>
    <td width="50" class="headerss" valign="bottom"><img src="/images/header_bg_l.gif" width="13" height="169" /></td>
    <td height="170" class="headerss" align="center" valign="top"><img src="/images/title.gif" width="413" height="165" /></td>
    <td width="50" class="headerss" valign="top"><img src="/images/header_bg_r.gif" width="52" height="170" /></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="content_table">
  <tr>
    <td width="50" class="content_l">&nbsp;</td>
    <td class="content"><ul class="tabs_ul">
      <li class="tabs_ul_li"><a href="/sp.asp?xdURL=webActivity/actPage1.asp&mp=<%=mp%>&id=<%=activityId%>">活動首頁</a></li>
      <li class="tabs_ul_li_active"><a href="#">活動說明</a></li>
    </ul>
      <p>1.活動時程：2011/11/21下午3時起至2011/11/30下午3時止。<br />
2.全國民眾凡具中華民國國籍，不限年齡、職業、性別，加入成為本網站會員後，皆可參加此活動。<a href="/sp.asp?xdURL=coamember/member_Join.asp&mp=1" class="link">馬上加入會員 </a><br />
3.參與會員請以中文姓名註冊，並填寫真實有效之個人資料與連絡方式，各會員請勿使用相同連絡信箱或臨時信箱註冊。若您已用英文拼音註冊，請於活動期間內，來信至km@mail.coa.gov.tw，確認您的身份，請管理員為您修改成中文名字。<br />
4.如有任何因電腦、網路、電話、技術或不可歸責於主辦單位之事由，而使參加者登錄之資料有延遲、遺失、錯誤、無法辨識或毀損之情況，主辦單位不負任何法律責任，參加者亦不得因此異議。<br />
5.為求公平，同一得獎者不拘獎項，限得獎一次，若有重複中獎情況，主辦單位得以保留最大獎項給得獎者，並取消較小獎項之得獎資格，以順位候補。若有得獎者會員之註冊信箱帳號重複，亦視同重複中獎情況處理。得獎者若使用臨時信箱或不正確信箱註冊，則視為得獎無效，以順位候補。<br />
6.正式得獎名單應以本網站公佈資料為準。得獎名單將於活動結束後2週內公佈於活動網站，並將以e-mail及電話方式(若有留電話號碼者)通知得獎者，中獎民眾須於期限內依通知進行身份確認，逾期、或身份不符者，皆視為放棄得獎資格，不另候補。各獎項將於公佈名單後1個月內陸續寄送。<br />
7.得獎者須正確填寫領獎收據，並附上身分證影本以便進行核對領獎身分。得獎者個人資料僅供本次核對領獎身分用，不予退還。<br />
8.活動得獎贈品僅限郵寄台、澎、金、馬地區。<br />
9.贈品寄出將再以e-mail與電話方式通知得獎者，未收到贈品者請於兩個月內提出詢問反應，逾期不受理亦不擔負賠償責任。<br />
10.參與者若有違反本競賽規則及本網站會員註冊同意書條款所列之規定，一經告發查證，主辦單位可逕行刪除活動資格及會員停權作業；該會員若為得獎者，則獎項將追回，採順位方式候補，不另行通知。相關法律刑責將由參賽者自行負責，若造成第三者之權益損失，參賽者得自負完全法律責任，不得異議。本網站亦將保留相關法律追訴之權利。<br />
11.活動期間為維護公平性，禁止網友從事外掛程式或任何影響活動公平之行為，一旦查證告發，主辦單位有權不需說明，直接刪除其抽獎與得獎資格，並可採順位候補不另行通知。<br />
12.本活動所涉及的個人資料蒐集、運用以及關於個人資料的保密，將適用農委會網站的隱私權保護政策。<br />
13.活動內容如遇不可抗拒之因素須進行變更，應以本網站公告為依據。主辦單位保留修改活動辦法、活動獎項及審核中獎資格之權利。</p></td>
    <td width="50" class="content_r">&nbsp;</td>
  </tr>
  <tr class="footer2">
    <td>&nbsp;</td>
    <td align="center"><img src="/images/footer.gif" width="318" height="150" /></td>
    <td>&nbsp;</td>
  </tr>
</table>

<%		
	end if
%>