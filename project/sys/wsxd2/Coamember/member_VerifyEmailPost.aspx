<!--#Include virtual = "/inc/client.inc" -->
<!--#include file = "des.inc" -->
<%
	on error resume next

	' Form collection elements
	Dim FcEmail, FcCaptcha, FcCheckCodeforMail
	Dim Err

	Err = false
	FcEmail = Request("email")
	FcCaptcha = Request("captcha")
	FcCheckCodeforMail = Request("CheckCodeforMail")
	
	' SQL injection check
	If InStr(FcEmail,"'")>0 Then
		Err = true
		response.write "<script>alert(' E-mail 格式錯誤！');history.back();</script>"
	End If

	' Captcha Verify
	If Not Err And (FcCaptcha <> FcCheckCodeforMail) Then
		Err = true
		response.write "<script>alert('圖片驗證碼錯誤，請重新填寫！');history.back();</script>"
	End If
	
	if Not Err Then
		
		' Recordset elements
		Dim RsCount
		Set rs2 = conn.Execute("SELECT COUNT(email) AS count FROM Member WHERE (email = '" & FcEmail & "')")	
		
		if not rs2.eof then
			RsCount = rs2("count")
		end if

		' Email Verify
		If (RsCount <> 0) Then
			Err = true
			response.write "<script>alert(' E-mail 已被登記使用，請重新輸入！');history.back();</script>"
		End If
		
	end if
%>
<%

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
		Response.Write("<BR>Reasonxx: " &  htPageDomCheck.parseError.reason)
		Response.End()
	end if

  	rtnVal = nullText(htPageDomCheck.selectSingleNode("SystemParameter/GIPconfig/" &funcName))
  	getGIPApConfigText = rtnVal
end function

Function Send_Email_authenticate(S_Email,R_Email,Re_Sbj,Re_Body, sendusername, sendpassword, SMTPauthenticate1)
	
	SMTPServer = getGIPApconfigText("EmailServerIp")
	SMTPServerPort = getGIPApconfigText("EmailServerPort")
	SMTPSsendusing = getGIPApconfigText("EmailServerSendUsing")

	Set objEmail = CreateObject("CDO.Message")
    objEmail.bodypart.Charset = "UTF-8"'response.charset
	objEmail.From       = S_Email
    objEmail.To         = R_Email
    objEmail.Subject    = Re_Sbj
    objEmail.HTMLbody   = Re_Body
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = SMTPSsendusing
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPServerPort
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = sendusername
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = sendpassword
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = SMTPauthenticate1

    objEmail.Configuration.Fields.Update
    objEmail.Send
    set objEmail = nothing

End Function
%>
<%
If Not Err Then

	Dim key, token, returnUrl, urlText, urlAction, EmailTo, EmailFrom, EmailTitle, Body
	
	key = session("WebDESKey")			'加解密金鑰(Len=14)
	Set DesCrypt = New Cls_DES
	token = DesCrypt.DES(FcEmail,key,0)	'加密
	Set DesCrypt = Nothing

	returnUrl = request("returnUrl")
	urlText = session("WebURL") 
	urlAction = session("WebURL") & "ConfirmEmail.asp?email=" & FcEmail & "&token=" & token & "&returnUrl=" & returnUrl
	EmailTo = FcEmail
	EmailFrom = getGIPApconfigText("EmailFrom")
	EmailTitle = "農業知識入口網註冊 Email 確認信"

	Body="<div>您好! 歡迎您加入農業知識入口網的行列，<br />"
	Body=Body & "這封信是由農業知識入口網的會員註冊系統所寄出，<br />"
	Body=Body & "請點選下面的網址來以確認 Email 並進行註冊的下一步驟：<br /><br />" 
	Body=Body & "<a href='" & urlAction & "'>" & urlText & "</a><br /><br />" 
	Body=Body & "※ 如果您不曾提出註冊申請，請您直接刪除本信，無須理會！<br />"
	Body=Body & "※ 此信為系統發生，請勿回覆<br /><br />"
	Body=Body & "農業知識入口網 敬上<br /></div>"
	
	dim SMTPusername
    dim SMTPpassword
    dim SMTPauthenticate

	SMTPusername = getGIPApconfigText("Emailsendusername")
	SMTPpassword = getGIPApconfigText("Emailsendpassword")
	SMTPauthenticate = getGIPApconfigText("Emailsmtpauthenticate")

	call Send_Email_authenticate(EmailFrom, EmailTo, EmailTitle, Body, SMTPusername, SMTPpassword, SMTPauthenticate)
%>
<div>
    <div class="path">
        目前位置：<a title="首頁" href="mp.asp">首頁</a>&gt; 加入會員</div>
    <h3>
        <b>農業知識入口網會員註冊系統</b></h3>
    <div id="Magazine">
        <div class="Event">
            <div class="experts">
                <div>
                    確認信已寄出，請收信！
                    <br />
	<%
	response.write("Email To: " + EmailTo + "<br>")
	response.write("Email From: " + EmailFrom + "<br>")
	response.write("Email Title: " + EmailTitle + "<br>")
	%>
                    <br />
                </div>
            </div>
        </div>
    </div>
</div>

<% End If %>
