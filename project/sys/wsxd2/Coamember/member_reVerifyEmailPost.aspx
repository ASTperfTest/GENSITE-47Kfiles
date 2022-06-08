<!--#Include virtual = "/inc/client.inc" -->
<!--#include file = "des.inc" -->
<%
	on error resume next
	
	' Form collection elements
	Dim FcAccount, FcPassword, FcEmail
	Dim Err

	Err = false
	FcEmail = Request("email")
	FcAccount = Request("account")
	FcPassword = Request("password")

	' SQL injection check
	If InStr(FcEmail,"'")>0 Then
		Err = true
		response.write "<script>alert(' E-mail 格式錯誤！');history.back();</script>"
	End If
	
	If FcAccount = "" Then
		Err = true
		response.write "<script>alert('尚未登入！');window.location.href='mp.asp';</script>"
	End If
	
	if Not Err then
		
		' Recordset elements
		Dim RsPasswd, RsEmail, RsCount

		Set rs1 = conn.Execute("SELECT passwd, email FROM Member WHERE (account = '" & FcAccount & "')")	
		
		if not rs1.eof then
			RsPasswd = Trim(rs1("passwd"))
			RsEmail = Trim(rs1("email"))
		end if

		Set rs2 = conn.Execute("SELECT COUNT(email) AS count FROM Member WHERE (account <> '" & FcAccount & "') AND (email = '" & FcEmail & "')")	
		
		if not rs2.eof then
			RsCount = rs2("count")
		end if
	
		' Password Verify
		If Not Err And (FcPassword <> RsPasswd) Then
			Err = true
			response.write "<script>alert('密碼錯誤，請重新輸入！');history.back();</script>"
		End If

		' Email Verify
		If Not Err And (RsCount > 0) Then
			Err = true
			response.write "<script>alert(' E-mail 已被登記使用，請重新輸入！');history.back();</script>"
		End If


		If Not Err And (RsCount = 0) Then
			conn.Execute("UPDATE Member SET ValidCount=ValidCount + 1  WHERE (account = '" & FcAccount & "')") 
		End If

		' Email Update DB
		If Not Err And (FcEmail <> RsEmail And RsCount = 0) Then
			conn.Execute("UPDATE Member SET email = '" & FcEmail & "' WHERE (account = '" & FcAccount & "')") 
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

	Dim key, token, updateMemID, urlText, urlAction, EmailTo, EmailFrom, EmailTitle, Body
	
	key = session("WebDESKey")			'加解密金鑰(Len=14)
	Set DesCrypt = New Cls_DES
	token = DesCrypt.DES(FcEmail,key,0)	'加密
	Set DesCrypt = Nothing

	updateMemID = FcAccount
	urlText = session("WebURL") 
	urlAction = session("WebURL") & "ConfirmEmail.asp?email=" & FcEmail & "&token=" & token & "&id=" & updateMemID
	EmailTo = FcEmail
	EmailFrom = getGIPApconfigText("EmailFrom")
	EmailTitle = "農業知識入口網註冊 E-mail 確認信"

	Body="<div>您好! 歡迎您加入農業知識入口網的行列，<br />"
	Body=Body & "這封信是由農業知識入口網的會員註冊系統所寄出，<br />"
	Body=Body & "請您點選信中的網址以確認您的 E-mail ：<br /><br />" 
	Body=Body & "<a href='" & urlAction & "'>確認Email</a><br /><br />" 
	Body=Body & "※ 如果您不曾提出註冊申請，請您直接刪除本信，無須理會！<br />"
	Body=Body & "※ 此信為系統自動發送，請勿回覆<br /><br />"
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
        目前位置：<a title="首頁" href="mp.asp">首頁</a>&gt; 重新認證 E-mail</div>
    <h3>
        <b>會員 E-mail 認證</b></h3>
    <div id="Magazine">
        <div class="Event">
            <div class="experts">
                <div>確認信已寄出，請您點選信中的網址以確認您的 E-mail 
                    <br />
	<%
	response.write("Email To: " + EmailTo + "<br>")
	'response.write("Email From: " + EmailFrom + "<br>")
	'response.write("Email Title: " + EmailTitle + "<br>")
	%>
                    <br />
                    <font color="red">(如果您一直沒有收到Email，可以重新認證並修改Email)</font><br/>
					<font color="red">(收信時請注意認證信是否位於<垃圾信箱>內)</font><br/>
					<font color="red">(若是沒收到認證信請連絡KM@mail.coa.gov.tw)</font><br/>
                </div>
            </div>
        </div>
    </div>
</div>

<% End If %>
