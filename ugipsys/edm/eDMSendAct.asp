<%@ CodePage = 65001 %>
<% 
	Response.Expires = 0
	response.charset="utf-8"
	HTProgCode = "DATAEDM"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%	
	Dim baseDSDId : baseDSDId = "7"
	Dim ctUnitId : ctUnitId = "2162"

	Dim SMTPServer : SMTPServer = "mail.coa.gov.tw"
	Dim SMTPServerPort : SMTPServerPort = 25
	Dim SMTPSsendusing : SMTPSsendusing = 2
	Dim SMTPusername : SMTPusername = "taft_km"
	Dim SMTPpassword : SMTPpassword = "P@ssw0rd"
	Dim SMTPauthenticate : SMTPauthenticate = "1"

	Dim epTreeID : epTreeID = "21"		
	Dim sender : sender = "km@mail.coa.gov.tw"
	
	Server.ScriptTimeOut = 2000
	on error resume next

	Function Send_Email_authenticate(S_Email,R_Email,Re_Sbj,Re_Body)
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
	    set objEmail = nothing
	End Function

	Dim sCount : sCount = 0
	Dim sTitle : sTitle = ""
	Dim xBody : xBody = ""
	sql = "SELECT TOP 1 * FROM CuDTGeneric WHERE iCTUnit = " & ctUnitId & " ORDER BY xPostDate DESC"
	Set rs = conn.execute(sql)
	If not rs.eof Then	
		sTitle = rs("sTitle")
		xBody = rs("xBody")
	end if
	rs.close
	Set rs = nothing
		
	sql = "SELECT email FROM Epaper WHERE CtRootID = " & epTreeID
		
	Set rs = conn.execute(sql)		
	while not rs.eof
		sCount = sCount + 1
		Call Send_Email_authenticate(sender, trim(rs("email")), sTitle, xBody)
		rs.movenext
	wend
	rs.close
	Set rs = nothing
%>
<script language=vbs>
	alert("eDM發送成功！共 <%=sCount%> 份")
	document.location.href = "eDMSend.asp"
</script>
