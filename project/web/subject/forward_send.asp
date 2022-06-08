<%@ CodePage = 65001 %><%
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
progPath="D:\hyweb\GENSITE\project\web\subject\forward_send.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("name", "message")
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
%>
<%
response.charset = "utf-8"
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1%>
<!--#include virtual = "/inc/checkURL.inc" -->
<?xml version="1.0"  encoding="utf-8" ?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
call CheckURL(Request.QueryString)
 function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

function getGIPApConfigText(byVal funcName)
    dim htPageDomCheck
    dim LoadXMLCheck
    
	set htPageDomCheck = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomCheck.async = false
	htPageDomCheck.setProperty("ServerHTTPRequest") = true
	if saveText <> "1" or saveText <> "Y" then
	    saveText = "N"
	end if

    LoadXMLCheck = "http://kmwebsys.coa.gov.tw/GenGipDSD/sysApPara.xml"
	xv = htPageDomCheck.load(LoadXMLCheck)
	if htPageDomCheck.parseError.reason <> "" then
		Response.Write("sysApPara.xml parseError on line " &  htPageDomCheck.parseError.line)
		'Response.Write("<BR>LoadXML: " &  LoadXML)
		Response.Write("<BR>Reasonxx: " &  htPageDomCheck.parseError.reason)
		Response.End()
	end if

  	rtnVal = nullText(htPageDomCheck.selectSingleNode("SystemParameter/GIPconfig/" &funcName))
  	getGIPApConfigText = rtnVal
end function
 
Function Send_Email(xxFrom,xxTo,xxSubject,xxBody)

SMTPServer = getGIPApconfigText("EmailServerIp")
SMTPServerPort = getGIPApconfigText("EmailServerPort")
SMTPSsendusing = getGIPApconfigText("EmailServerSendUsing")
SMTPusername = getGIPApconfigText("Emailsendusername")
SMTPpassword = getGIPApconfigText("Emailsendpassword")
SMTPauthenticate = getGIPApconfigText("Emailsmtpauthenticate")    

	Set objEmail = CreateObject("CDO.Message")
    objEmail.bodypart.Charset = "UTF-8"'response.charset
	objEmail.From       = xxFrom
    objEmail.To         = xxTo
    objEmail.Subject    = xxSubject
    objEmail.HTMLbody   = xxBody
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = SMTPSsendusing
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPServerPort
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPusername
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPpassword
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") =SMTPauthenticate

    objEmail.Configuration.Fields.Update
    objEmail.Send
    set objEmail=nothing


End Function
%> 

<%

xFrom = request("email2")
SEmailName = request("name")
SEmail = "test@test.com"
if xFrom ="" then xFrom =""" "& SEmailName &" "" <"& SEmail &"> "


xTo = request("email")
if xTo = "" then xTo ="test@test.com"

message = request("message")


xctUrl = request("ctUrl")
if xctUrl = "" then xctUrl="xItem=6&ctNode=109&mp=47"

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
	qStr = xctUrl
	if instr(qStr,"mp=")=0 then qStr = qStr & "&mp=" & session("mpTree")
	
	xv = oxml.load(session("myXDURL") & "/wsxd/xdcp.asp?" & qStr)

  if oxml.parseError.reason <> "" then 
    Response.Write("htPageDom parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if
  	xmyStyle= nullText(oxml.selectSingleNode("//hpMain/MpStyle"))
	if xmyStyle="" then	
		xmyStyle=nullText(oxml.selectSingleNode("//hpMain/myStyle"))
  		if xmyStyle="" then xmyStyle=session("myStyle")
	end if	
	xSubject = nullText(oxml.selectSingleNode("//hpMain/MainArticle/Caption"))

	fStyle = oxml.selectSingleNode("//xslData").text		
	oxsl.load(server.mappath("xslGip/" & xmyStyle & "/cp.xsl"))
	
  	if fStyle <> "" then
		set fxsl = server.createObject("microsoft.XMLDOM")
		fxsl.load(server.mappath("xslGip/" & xmyStyle & "/" & fStyle & ".xsl"))		
		set oxRoot = oxsl.selectSingleNode("xsl:stylesheet")
	
	on error resume next
		for each xs in fxsl.selectNodes("//xsl:template")
			set nx = xs.cloneNode(true)
			ckStr = "@match='" & nx.getAttribute("match") & "'"
			if nx.getAttribute("mode")<>"" then		ckStr = ckStr & " and @mode='" & nx.getAttribute("mode") & "'"
			set orgEx = oxRoot.selectSingleNode("//xsl:template[" & ckStr & "]")
			oxRoot.removechild orgEx
			oxRoot.appendChild nx
		next
		for each xs in fxsl.selectNodes("//msxsl:script")
			set nx = xs.cloneNode(true)
			oxRoot.appendChild nx
		next
	end if 
	
	Response.ContentType = "text/HTML" 
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">","")
	outString =  replace(outString,"&amp;","&")
	outString = replace(outString,"href=""","href="""& session("myWWWSite"))
	outString = replace(outString,"href="""& session("myWWWSite")&"mailto:","href=""mailto:")
	outString = replace(outString,"src=""","src="""& session("myWWWSite"))
	outString = replace(outString,"subject//","")
	'xBody = "您的朋友：" & SEmailName & "   寄給你一封訊息 <br/>"
	'xBody = xBody & "-------------------------------------------------------"
	'xBody = xBody & "<br/>給您的留言：<br/>"
	'xBody = xBody & message  &"<br/>"
	
	xBody = xBody & outString
	'response.write xBody
	call Send_Email (xFrom,xTo,xSubject,xBody)
	msg_desc = "alert('轉寄成功'); self.close();"

%>
<body bgcolor="#FFFFFF" onload="<%=msg_desc%>">
</body>