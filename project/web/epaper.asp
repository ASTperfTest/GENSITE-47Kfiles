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
progPath="D:\hyweb\GENSITE\project\web\epaper.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("eid")
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
	Response.Expires = 0
	HTProgCode="GW1M51"
	HTProgPrefix="ePub" 
%>
<!--#Include virtual = "/inc/dbUtil.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/client.inc" -->
<!--#include file = "inc/web.sqlInjection.inc" -->
<%
	function nullText(xNode)
	  on error resume next
	  xstr = ""
	  xstr = xNode.text
	  nullText = xStr
	end function

	epaperID = pkStrWithSriptHTML( request("eid"), "" )	
	
	if epaperID = "null" then
		Response.Write("您選擇的電子報不存在")
 		Response.End()
	end if	
	epaperID = request("eid")
	if not isnumeric(epaperID) then
		Response.Write("電子報錯誤")
 		Response.End()
	end if	

	'----Load epaper.xml
	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty "ServerHTTPRequest", true
	'LoadXML = server.mappath("/site/" & session("mySiteID") & "/public/ePaper/" & formID & ".xml")
	'LoadXML = server.mappath("/site/coa/public/ePaper/ep10.xml")
	'LoadXML = ("http://kmbeta.hyweb.com.tw/public/ePaper/ep" & epaperID & ".xml")
	LoadXML = server.mappath("/public/ePaper/ep" & epaperID & ".xml")

	xv = oxml.load(LoadXML)

  if oxml.parseError.reason <> "" then
  	Response.Write("XML parseError on line " &  oxml.parseError.line)
   	Response.Write("<BR>Reasonaa: " &  oxml.parseError.reason)
    	Response.End()
  end if
  
  '----Load epaper.xsl
	set oxsl = server.createObject("microsoft.XMLDOM")
	'oxsl.load(server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/ePaper"&epTreeID &".xsl"))
	'oxsl.load(server.MapPath("/site/coa/public/ePaper/ePaper143.xsl"))	
	sqlCom = "SELECT * FROM EpPub WHERE epubId=" & pkStr(epaperID,"")
	Set RSmaster = conn.execute(sqlcom)
	'set xsl type
	ePubType = RSmaster("pubType")
	oxsl.load(server.MapPath("/site/coa/public/ePaper/ePaper" & ePubType & ".xsl"))	
	'oxsl.load("ePaper29.xsl")
	
	Response.ContentType = "text/HTML" 
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	outString = replace(outString,"&amp;","&")
	response.write replace(outString,"&amp;","&")

	response.end

%>
