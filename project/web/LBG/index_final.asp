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
progPath="D:\hyweb\GENSITE\project\web\LBG\index_final.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array()
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
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-TW" lang="zh-TW">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>「虛擬種植，知識升值」- 豆仔花園</title>
<script language="javascript"> AC_FL_RunContent = 0; </script>
<script language="javascript"> DetectFlashVer = 0; </script>
<script src="AC_RunActiveContent.js" language="javascript"></script>
<script language="JavaScript" type="text/javascript">
<!--
// -----------------------------------------------------------------------------
// 全域
// 需要 Flash 的主要版本
var requiredMajorVersion = 9;
// 需要 Flash 的次要版本
var requiredMinorVersion = 0;
// 需要的 Flash 版本
var requiredRevision = 45;
// -----------------------------------------------------------------------------
// -->
</script>
</head>
<style type="text/css">
<!--
body {
	background-image: url(bg001.gif);
}
#bottom p {
	margin: 2px 0px 0px;
	font-size: 12px;
	color: #333333;
	padding: 0px;
	float: left;
	line-height: 18px;
	text-align: left;
}
#bottom img {
	margin-left: 100px;
	float: left;
	margin-right: 10px;
	padding-bottom: 20px;
}
#bottom {
	width: 760px;
	margin-top: 10px;
}
#bottom a {
	font-size: 12px;
	color: #006666;
}
-->
</style></head>
<body bgcolor="#ffffff">

<!--影片中使用的 URL-->
<!--影片中使用的文本-->

<center>
<script language="JavaScript" type="text/javascript">
<!--
if (AC_FL_RunContent == 0 || DetectFlashVer == 0) {
	alert("這個頁面必須具備 AC_RunActiveContent.js。");
} else {
	var hasRightVersion = DetectFlashVer(requiredMajorVersion, requiredMinorVersion, requiredRevision);
	if(hasRightVersion) {  // 如果偵測到可接受的版本
		// 嵌入 flash 影片
		AC_FL_RunContent(
			'codebase', 'http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,45,0',
			'width', '720',
			'height', '420',
			'src', 'LITTLE_BEAN',
			'quality', 'high',
			'pluginspage', 'http://www.macromedia.com/go/getflashplayer',
			'align', 'middle',
			'play', 'true',
			'loop', 'false',
			'scale', 'showall',
			'wmode', 'window',
			'devicefont', 'false',
			'id', 'LITTLE_BEAN',
			'bgcolor', '#ffffff',
			'name', 'LITTLE_BEAN',
			'menu', 'true',
			'allowNetworking','all',
			'allowScriptAccess','always',
			'allowFullScreen','false',
			'movie', 'LITTLE_BEAN',
			'salign', '' ,
			'flashvars' ,'hostIP=<%=request.servervariables("SERVER_NAME") %>'  
			); //end AC code
	} else {  // flash 太舊或我們無法偵測到外掛程式
		var alternateContent = '替代的 HTML 內容應置於此處。 '
			+ '這個內容需要 Adobe Flash Player。 '
			+ '<a href=http://www.macromedia.com/go/getflash/>取得 Flash</a>';
		document.write(alternateContent);  // 插入不是 Flash 的內容
	}
}
// -->
</script>

<!--影片中用到 URL-->
<!--影片中用到文字-->

  <div id="bottom">
	<span style="color:red;font-size: 12px;">本遊戲排行榜每一小時更新一次</span><br/>
	<a href="http://www.macromedia.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash"><img src="get_flash_player.gif" alt="FLASH PLAYER連接按鈕" border="0" /></a>
	<p>請使用IE 6.0以上 最佳觀看解析度1024x768．桌面設定以系統預設字型為佳<br/>
	本網站之著作權為行政院農業委員會所有，非經同意禁止轉載<br/>
	
	</p>
  </div>
	
</center>

<script type="text/javascript">
var gaJsHost=(("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try{
var pageTracker=_gat._getTracker("UA-9195501-1");
pageTracker._trackPageview();
}catch(err)
{}
</script>

</body>

</html>
