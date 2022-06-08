<!--#include virtual = "/inc/checkGIPAPconfig.inc" -->
<%

Sub SendEmail( questionId, expertId, mailBody, myrand, icuitem )
	
	'S_Email = "km@mail.coa.gov.tw"
	S_Email = """"&getGIPApconfigText("EmailFromName")&""" <"&getGIPApconfigText("EmailFrom")&">"
	'S_Email = "vincent@hyweb.com.tw"
	R_Email = GetExpertEmail( expertId )
	
	sTitle = "農業知識家專家補充邀請：「" & GetsTitle( questionId, "") & "」"
		
	xBody = GetEmailBodyFromFile()
	xBody = ReplaceTag(xBody , mailBody, expertId, questionId, myrand, icuitem)

	Call Send_Email( S_Email, R_Email, sTitle, xBody)

End Sub

Function Send_Email( S_Email, R_Email, Re_Sbj, Re_Body )
    dim SMTPusername
    dim SMTPpassword
    dim SMTPauthenticate

	SMTPusername = getGIPApconfigText("Emailsendusername")
	SMTPpassword = getGIPApconfigText("Emailsendpassword")
	SMTPauthenticate = getGIPApconfigText("Emailsmtpauthenticate")

    'SMTPusername = "km"
    'SMTPpassword = ""
    'SMTPusername = "testemail"
    'SMTPpassword = "1q2w3e4r"
    'SMTPauthenticate = "1"

    call Send_Email_authenticate(S_Email,R_Email,Re_Sbj,Re_Body, SMTPusername, SMTPpassword, SMTPauthenticate)
End Function

Function Send_Email_authenticate(S_Email,R_Email,Re_Sbj,Re_Body, SMTPusername, SMTPpassword, SMTPauthenticate)
		dim SMTPServer
		dim SMTPServerPort
		dim SMTPSsendusing

		SMTPServer = getGIPApconfigText("EmailServerIp")
		SMTPServerPort = getGIPApconfigText("EmailServerPort")
		SMTPSsendusing = getGIPApconfigText("EmailServerSendUsing")		

		'SMTPServer = "mail.coa.gov.tw"    
		'SMTPServer = "mail.hyweb.com.tw"
		'SMTPServerPort = 25
		'SMTPSsendusing = 2
		
    Set objEmail = CreateObject("CDO.Message")
    objEmail.bodypart.Charset = "UTF-8"
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
%>