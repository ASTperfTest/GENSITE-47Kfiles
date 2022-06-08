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
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\webActivity\actPage1.asp"

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
	width:460px;
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
    <td width="50" class="content_l"><img src="/images/blank.gif" /></td>
    <td class="content"><ul class="tabs_ul">
      <li class="tabs_ul_li_active"><a href="#">活動首頁</a></li>
      <li class="tabs_ul_li"><a href="/sp.asp?xdURL=webActivity/actPage2.asp&mp=<%=mp%>&id=<%=activityId%>">活動說明</a></li>
    </ul>
      <p>農業知識入口網為能持續提供給眾多使用者更豐富的內容與友善便利的系統功能，自2011/11/21下午4時起至2011/11/30下午4時止，只要您登入會員，並填寫完成問卷調查內容，就有機會抽中獎項包括：這禮日月潭禮盒 4名、JetFlash600隨身碟16GB 8名及加比山女兒禮盒 13名，另外還提供限量豆仔玉米杯20名給會員及統一1,000禮卷3名給創意回覆意見的會員，共48個得獎名額。</p>
      <p class="crnter">好康&quot;豆&quot;相報，歡迎大家一起來參與。</p>
      <div id="bt"><a href="/sp.asp?xdURL=webActivity/actPage5.asp&mp=<%=mp%>&id=<%=activityId%>"><img src="/images/bt.jpg" width="181" height="63" /></a></div>
      <div id="present">
        <h1>活動獎項 </h1>
        <table width="400" border="0" align="center" cellpadding="0" cellspacing="0" id="present_list">
          <tr>
            <th width="25%">獎項</th>
            <th width="60%">獎品內容</th>
            <th width="15%">數量</th>
          </tr>
          <tr>
            <td align="center">特獎</td>
            <td align="center"><a target="_blank" href="http://www.taiwansoap.com.tw/m_product/product_da.php?eProductcy_id=62&eProductcy_sid=2709996464&eProductda_id=112">這禮日月潭禮盒</a></td>
            <td align="center">4</td>
          </tr>
          <tr>
            <td align="center">普獎</td>
            <td align="center"><a target="_blank" href="http://shop.transcend.com.tw/product/ItemDetail.asp?ItemID=TS16GJF600">JetFlash600隨身碟16GB</a></td>
            <td align="center">8</td>
          </tr>
          <tr>
            <td align="center">感謝獎</td>
            <td align="center"><a target="_blank" href="http://www.farms.tw/shop/goods.php?id=989">加比山女兒禮盒</a></td>
            <td align="center">13</td>
          </tr>
          <tr>
            <td align="center">肯定獎</td>
            <td align="center">限量豆仔玉米杯</td>
            <td align="center">20</td>
          </tr>
          <tr>
            <td align="center">創意回覆獎</td>
            <td align="center">統一禮卷1,000元</td>
            <td align="center">3</td>
          </tr>
          <tr>
            <td colspan="2" align="center">總計</td>
            <td align="center">48</td>
          </tr>
        </table>
        <ul>
          <li>活動贈品無法由得獎者指定顏色 </li>
          <li>贈品相關規格請點選連結查看，詳細介紹請自行洽詢該產品官網 </li>
          <li>主辦單位保留修改活動辦法、活動獎項及審核中獎資格之權利 </li>
        </ul>
      </div></td>
    <td width="50" class="content_r"><img src="/images/blank.gif" /></td>
  </tr>
  <tr class="footer2">
    <td>&nbsp;</td>
    <td align="center"><img src="/images/footer.gif" width="318px" height="150px" /></td>
    <td>&nbsp;</td>
  </tr>
</table>
</div>

<%		
	end if
%>