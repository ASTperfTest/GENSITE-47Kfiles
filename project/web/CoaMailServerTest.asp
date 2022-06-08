<%
	Server.ScriptTimeOut = 1800
	Response.Buffer = True
	Response.Charset = "big5"

	dim SMTPServer
	dim SMTPServerPort
	dim SMTPSsendusing

	dim SMTPusername
	dim SMTPpassword
	dim SMTPauthenticate

    
  SMTPServer = "mail.coa.gov.tw"
  SMTPServerPort = 25    
  SMTPSsendusing = 2

  emailBody     = "TEST,TEST,TEST"
  emailSubject  = "TEST"
    
	emailFrom     = "taft_km@mail.coa.gov.tw"
  emailTo       = "vivian.shr@gmail.com"
		
  Call Send_Email(emailFrom,emailTo,emailSubject,emailBody)
%>

<%
Function Send_Email(S_Email,R_Email,Re_Sbj,Re_Body)
    dim SMTPusername
    dim SMTPpassword
    dim SMTPauthenticate

		SMTPusername = "taft_km"
		SMTPpassword = "P@ssw0rd"		
    SMTPauthenticate = "1"

    call Send_Email_authenticate(S_Email,R_Email,Re_Sbj,Re_Body, SMTPusername, SMTPpassword, SMTPauthenticate)
End Function

Function Send_Email_authenticate(S_Email,R_Email,Re_Sbj,Re_Body, SMTPusername, SMTPpassword, SMTPauthenticate)
    Set objEmail = CreateObject("CDO.Message")
    objEmail.bodypart.Charset = response.charset
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
		response.write "done"
End Function

%>