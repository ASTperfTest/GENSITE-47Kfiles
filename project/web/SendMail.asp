

<html>
<head><meta content="text/html; charset=UTF-8" http-equiv="Content-Type"></head>
<body>
<%

mailType = request("type")
if mailType <> "2" then mailType = "1"

select case mailType
    case "1"
        ePaperTitle="農業知識入口網-系統問題反應"
    case else
        ePaperTitle="農業知識入口網-檢舉信"
end select



function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

function checkGIPApConfig(funcName)
    dim htPageDomCheck
    dim LoadXMLCheck
    
	set htPageDomCheck = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomCheck.async = false
	htPageDomCheck.setProperty("ServerHTTPRequest") = true

    LoadXMLCheck = session("sendMailConfigFile")
	xv = htPageDomCheck.load(LoadXMLCheck)
	if htPageDomCheck.parseError.reason <> "" then
		Response.Write("sysApPara.xml parseError on line " &  htPageDomCheck.parseError.line)
		'Response.Write("<BR>LoadXML: " &  LoadXML)
		Response.Write("<BR>Reasonxx: " &  htPageDomCheck.parseError.reason)
		Response.End()
	end if

  	checkGIPApConfig = false
  	if UCASE(nullText(htPageDomCheck.selectSingleNode("SystemParameter/GIPconfig/" &funcName)))="Y" then
		checkGIPApConfig = true
    end if
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

    LoadXMLCheck = session("sendMailConfigFile")
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


'// save setting item text from sysApPara.xml
'// modify date: 2006/09/20
function saveGIPAPconfig(funcName, byVal saveText)
    dim rtnVal          '// return value
    dim htPageDomCheck
    dim LoadXMLCheck
    dim xvCheck
	rtnVal = ""
	
	set htPageDomCheck = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomCheck.async = false
	htPageDomCheck.setProperty("ServerHTTPRequest") = true

	LoadXMLCheck = session("sendMailConfigFile")
	
	xvCheck = htPageDomCheck.load(LoadXMLCheck)
	if htPageDomCheck.parseError.reason <> "" then
		Response.Write("sysApPara.xml parseError on line " &  htPageDomCheck.parseError.line)
		'Response.Write("<BR>LoadXML: " &  LoadXML)
		Response.Write("<BR>Reasonxx: " &  htPageDomCheck.parseError.reason)
		Response.End()
	end if

  	htPageDomCheck.selectSingleNode("SystemParameter/GIPconfig/" &funcName).text = saveText
  	rtnVal = htPageDomCheck.save(LoadXMLCheck)
  	saveGIPApConfig = rtnVal
end function


'// get setting item text from sysApPara.xml
'// modify date: 2006/01/06
function saveGIPAPconfigText(funcName, saveText)
    dim rtnVal          '// return value
    dim htPageDomCheck
    dim LoadXMLCheck
    dim xvCheck
	rtnVal = ""
	
	set htPageDomCheck = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomCheck.async = false
	htPageDomCheck.setProperty("ServerHTTPRequest") = true

	LoadXMLCheck = session("sendMailConfigFile")
	xvCheck = htPageDomCheck.load(LoadXMLCheck)
	if htPageDomCheck.parseError.reason <> "" then
		Response.Write("sysApPara.xml parseError on line " &  htPageDomCheck.parseError.line)
		'Response.Write("<BR>LoadXML: " &  LoadXML)
		Response.Write("<BR>Reasonxx: " &  htPageDomCheck.parseError.reason)
		Response.End()
	end if

  	htPageDomCheck.selectSingleNode("SystemParameter/GIPconfig/" &funcName).text = saveText
  	
  	rtnVal = htPageDomCheck.save(LoadXMLCheck)
  	saveGIPApConfigText = rtnVal
end function


%>
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/client.inc" -->
<!--#include file = "inc/web.sqlInjection.inc" -->
<% 	if Session("memID")="" then%>
<script language="javascript">
	alert("請先登入");
	window.close();	
</script>
<% 	response.end
    else %>
	<%
	    session("SendMail_txtDisCussion") = request("txtDisCussion")    '視窗關閉後保留輸入內容
		if (Request("CheckCodeforMail") <> Session("CheckCodeforMail")) then
	%>
	    <script language="javascript">
		    alert("圖片驗證碼錯誤，請重新填寫");		
		    window.close();
	    </script>	  	
<%
        response.End
		else
	        session("SendMail_txtDisCussion") =""
			txtDisCussion = request("txtDisCussion")
			txtDisCussion=Replace(txtDisCussion, vbCrLF, " <br />")
			url=request("urlRecord")
'response.write "ARTICLE_ID:" & request("ARTICLE_ID")
'response.end
'SMTPServer = "mail.coa.gov.tw"
'SMTPServerPort = 25
'SMTPSsendusing = 1
'SMTPSsendusing = 2



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
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") =SMTPauthenticate1

    objEmail.Configuration.Fields.Update
    objEmail.Send
    set objEmail=nothing

End Function

articleId = ""
if request("ARTICLE_ID") <> "" then
	articleId = request("ARTICLE_ID")
else
	articleId = request.form("ARTICLE_ID")
End If



sql ="INSERT INTO [mGIPcoanew].[dbo].[KNOWLEDGE_REPORT] "
sql = sql & "(ARTICLE_ID,CREATOR,CREATION_DATETIME,[TYPE],SOURCE_URL,DESCRIPTION,RESPONSE,LAST_MODIFIER,LAST_MODIFY_DATETIME,STATUS)"
sql= sql & " VALUES"
if mailType = "2" then    
    tmpTxtDisCussion=txtDisCussion & "<hr size=''1''/>" & "檢舉對象：" & replace(request("account"),"'","''") & "　" & replace(request("nickname") ,"'","''")
    sql= sql & " (" & articleId & ",'" & session("memID") & "', GETDATE() , " & mailType & ",'" & url & "','" & tmpTxtDisCussion & "','','" & session("memID") & "',GETDATE(),0)"
else
    sql= sql & " (" & articleId & ",'" & session("memID") & "', GETDATE() , " & mailType & ",'" & url & "','" & txtDisCussion & "','','" & session("memID") & "',GETDATE(),0)"
end if

'response.write sql
'response.end
conn.execute sql

'問題單回應，http://gssjira.gss.com.tw/browse/COAKM-49，寄件者email要讀取user註冊的user
sql2="select email from  dbo.Member where account='" & session("memID") &"'"
set rs = conn.Execute(sql2)

S_Email=rs("email")
R_Email=getGIPApconfigText("EmailFrom")

Body="帳號：" & session("memID") & "<br />"
Body=Body & "姓名：" & session("memName") & "<br />"
Body=Body & "暱稱：" & session("memNickName") & "<br />" 
Body=Body & "email：" & S_Email & "<br />" 

if mailType = "2" then
    Body=Body & "檢舉對象：" & request("account") & "　" & request("nickname") & "<br />"
end if

Body=Body & "文件來源：<a href='" & url & "'>查看</a><br />"

Body=Body & "問題描述：<br />" & txtDisCussion
	dim SMTPusername
    dim SMTPpassword
    dim SMTPauthenticate

    'SMTPusername = "taft_km"
    'SMTPpassword = "P@ssw0rd"
    'SMTPauthenticate = "1"
	
	SMTPusername = getGIPApconfigText("Emailsendusername")
	SMTPpassword = getGIPApconfigText("Emailsendpassword")
	SMTPauthenticate = getGIPApconfigText("Emailsmtpauthenticate")
	call Send_Email_authenticate(R_Email,R_Email,ePaperTitle,Body, SMTPusername, SMTPpassword, SMTPauthenticate)

%>


<script language="javascript">
	//alert(document.referrer);
	alert("已通知管理員");
	window.close();	
	
	//window.location.href=document.referrer;	
	//history.back();
	
</script>
<%
		end if	
	end if
%>

</body></html>