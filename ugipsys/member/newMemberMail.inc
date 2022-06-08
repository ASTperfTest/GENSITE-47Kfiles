<!--#include virtual = "/inc/checkGIPAPconfig.inc" -->
<%
dim SMTPServer
dim SMTPServerPort
dim SMTPSsendusing

dim SMTPusername
dim SMTPpassword
dim SMTPauthenticate

'SMTPServer = "mail.coa.gov.tw"
'SMTPServerPort = 25
'SMTPSsendusing = 1
'SMTPSsendusing = 2
'讀sysApPara.xml 裡的設定值
SMTPServer = getGIPApconfigText("EmailServerIp")
SMTPServerPort = getGIPApconfigText("EmailServerPort")
SMTPSsendusing = getGIPApconfigText("EmailServerSendUsing")

Function Send_Email_authenticate(S_Email,R_Email,Re_Sbj,Re_Body, SMTPusername, SMTPpassword, SMTPauthenticate)
    Set objEmail = CreateObject("CDO.Message")
    objEmail.bodypart.Charset = "UTF-8"'response.charset
    objEmail.From       = S_Email
    objEmail.To         = R_Email
    objEmail.Subject    = Re_Sbj
    objEmail.HTMLbody   = Re_Body
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = SMTPSsendusing
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPServerPort

    if SMTPusername <> "" then
        objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPusername
        objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPpassword
        objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = SMTPauthenticate
    end if


    objEmail.Configuration.Fields.Update
    objEmail.Send
    set objEmail=nothing
End Function

Function Send_Email(S_Email,R_Email,Re_Sbj,Re_Body)
    dim SMTPusername
    dim SMTPpassword
    dim SMTPauthenticate

    'SMTPusername = "taft_km"
    'SMTPpassword = "P@ssw0rd"
	'SMTPauthenticate = "1"
	
	SMTPusername = getGIPApconfigText("Emailsendusername")
	SMTPpassword = getGIPApconfigText("Emailsendpassword")
	SMTPauthenticate = getGIPApconfigText("Emailsmtpauthenticate")

    call Send_Email_authenticate(S_Email,R_Email,Re_Sbj,Re_Body, SMTPusername, SMTPpassword, SMTPauthenticate)
End Function



%>