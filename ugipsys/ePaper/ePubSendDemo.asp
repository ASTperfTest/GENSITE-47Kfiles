<%@ CodePage = 65001 %>
<% 
	Response.Expires = 0
	response.charset="utf-8"
	HTProgCode="GW1M51"
	HTProgPrefix="ePub" 
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<!--#include virtual = "/inc/checkGIPAPconfig.inc" -->
<%
Server.ScriptTimeOut = 5000
on error resume next

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

SMTPServer = getGIPApconfigText("EmailServerIp")
SMTPServerPort = getGIPApconfigText("EmailServerPort")
SMTPSsendusing = getGIPApconfigText("EmailServerSendUsing")

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

'Function Send_Email (S_Email,R_Email,Re_Sbj,Re_Body)
'  Set objNewMail = CreateObject("CDONTS.NewMail") 
'  objNewMail.MailFormat = 0
'  objNewMail.BodyFormat = 0 
	'   response.write S_Email&"[]"&R_Email&"[]"&Re_Sbj&"[]"&Re_Body
	'   response.write "<hr>"
	'   response.write Re_Body
	'   response.end
'  call objNewMail.Send(S_Email,R_Email,Re_Sbj,Re_Body)
'  Set objNewMail = Nothing
'End Function

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function
 
 	epTreeID = session("epTreeID")		'-- 電子報的 tree

	ePubID = request.queryString("ePubID")
	ePubMail = request.queryString("ePubMail")
	formID = "ep" & ePubID
	sqlCom = "SELECT * FROM EpPub WHERE epubId=" & pkStr(ePubID,"")
	Set RSmaster = Conn.execute(sqlcom)
	'set xsl type
	ePubType = RSmaster("pubType")
	
	'----Load epaper.xml
	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/site/" & session("mySiteID") & "/public/ePaper/" & formID & ".xml")
	
	'----檢查xml檔是否存在
  Set fso = server.CreateObject("Scripting.FileSystemObject")
  if not fso.FileExists(LoadXML) then
%>
		<script language=vbs>	
			alert "電子報內容檔不存在, 請先產生內容檔!"
			window.history.back	
		</script>		
<%	
	end if
	xv = oxml.load(LoadXML)
  if oxml.parseError.reason <> "" then
  	Response.Write("XML parseError on line " &  oxml.parseError.line)
  	Response.Write("<BR>Reason: " &  oxml.parseError.reason)
  	Response.End()
  end if  
  set dXml = oxml.selectSingleNode("ePaper")
  	
  '----Load epaper.xsl
	set oxsl = server.createObject("microsoft.XMLDOM")
	
	'response.write server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/ePaper"&epTreeID &".xsl")
	'response.end
	
	oxsl.load(server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/ePaper"&ePubType &".xsl"))
	
	'response.write "<XMP>"+eXml.xml+"</XMP>"

  '----參數處理
	ePaperTitle = dXml.selectSingleNode("ePaperTitle").text
	ePaperSEmail = dXml.selectSingleNode("ePaperSEmail").text
	ePaperSEmailName = dXml.selectSingleNode("ePaperSEmailName").text	
  S_Email=""""&ePaperSEmailName&""" <"&ePaperSEmail&">"
	sCount = 0
	set eXml = oxml.cloneNode(true)	
	Response.ContentType = "text/HTML" 
	outString = replace(exml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	outString = replace(outString,"&amp;","&")
	xmBody = replace(outString,"&amp;","&")

	Email_body=xmBody
	'response.end

'---- 3. 送 只訂閱 ePaper 的 mail ---------------------------------------------------------------------
	
	if ePubMail <> "" then
            R_Email=ePubMail
            sCount = sCount + 1
            Call Send_Email(S_Email,R_Email,ePaperTitle,Email_body)
	end If
If session("epTreeID") = "21" Then

	sql = "select title from EpPub where CtRootId = 21 AND ePubID = " & epubid
	set eprs = conn.Execute(sql)
	eptitle = eprs(0)

	'---vincent : insert into history epaper-20080317---

End If	

%>
	<script language=vbs>	
		alert "電子報發送完成!  共發送<%=sCount%>份電子報"
		window.close 	
	</script>		
	
