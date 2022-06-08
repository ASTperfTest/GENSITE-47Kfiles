<%@ CodePage = 65001 %>
<%response.expires = 0%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<title>忘記帳號密碼</title>
	<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>

<!-- #INCLUDE FILE="inc/dbutil.inc" -->
<%
	'要檢查的 request 變數名稱
	pArray = array("Email")
	for each str in pArray
		p = request(str)
		'要檢查的 pattern
		patern = array("<", ">", "%3C", "%3E", ";", "%27", "'")
	
		for each ptn in patern
			if (Instr(p, ptn) > 0) then
				response.end
			end if
		next
	next
%>

<%
	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
'Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
conn.ConnectionTimeout=0
conn.CursorLocation = 3
conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	function nullText(xNode)
	  on error resume next
	  xstr = ""
	  xstr = xNode.text
	  nullText = xStr
	end function

	Function Send_Email (S_Email,R_Email,Re_Sbj,Re_Body)
   Set objNewMail = CreateObject("CDONTS.NewMail") 
   objNewMail.MailFormat = 0
   objNewMail.BodyFormat = 0 
   call objNewMail.Send(S_Email,R_Email,Re_Sbj,Re_Body)
   Set objNewMail = Nothing
	End Function

	if request("task")="查詢" then
   	sql = "Select * from InfoUser where Email=" & pkStrWithSriptHTML(request("Email"), "")
   	set rs = conn.execute(sql)
		if rs.eof then
%>
			<script language=Javascript>
				alert ("查無此E-Mail!");
			</script>
<%
	else
		
		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysPara.xml"		
		xv = htPageDom.load(LoadXML)
		if htPageDom.parseError.reason <> "" then 
    	Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    	Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    	Response.End()
  	end if

		formCCMail = nullText(htPageDom.selectSingleNode("//DsdXMLEmail"))

		S_Email="""GIP 後台管理系統"" <gipService@hyweb.com.tw>"
		if formCCMail <> "" then _
			S_Email="""GIP 後台管理系統"" <" & formCCMail & ">"
      R_Email=rs("Email")
      Email_body="【 " & stripHTML(rs("UserName")) & " 】您好:" & "<br>" & _
      "下列為您的帳號/密碼" & _
			"<pre>      帳號: 【 " & stripHTML(rs("UserID")) & " 】" & "</pre>" & _
			"<pre>      密碼: 【 " & stripHTML(rs("Password")) & " 】" & "</pre>" & _
			"請查收." & "<br>"
      Call Send_Email(S_Email,R_Email,"GIP 帳號密碼通知",Email_body)
			'------------
%>
			<script language=Javascript>
	               alert (" 您的帳號及密碼, 我們將以 Email 方式通知您 !");
			window.close();
		</script>
<%
		end if
	end if
%>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<table border="0" width="300" cellspacing="1" cellpadding="0" class=body1>
  <tr> 
    <td width="90%" valign="middle">
      <p align="center">忘記帳號密碼查詢</p>
 </td>
</tr>    
  </table> 
  <table border="0" width="300" cellspacing="1" cellpadding="0" class=bg > 
    <tr> 
      <td width="100%" class=body> 
    <form name=fmReID method="POST" action="" onsubmit="Javascript: return datacheck();">  
        <p align="left"><br> 
        <div align="center">
<table border="1">
    <tr><th align=right>請輸入您的E-Mail:
    <tr><td align=right><input name="Email" size="20">
</table><BR>
<center>
<input type=SUBMIT name=task class=cbutton value="查詢">
<input type=button name=bye class=cbutton onclick="Javascript:byenow();" value="回首頁">          
    </form>      
      </td> 
    </tr>
  </table> 
</body>
</html>
<script language=VBS>
sub window_onLoad
    fmReID.Email.focus
end sub
</script>
<script language=javascript>
function byenow() {
	window.close();
}
function datacheck()
{    
  if (document.fmReID.Email.value=="")
  {
     alert("請輸入您的E-Mail!");
     return false;
  } 
}
</script>
