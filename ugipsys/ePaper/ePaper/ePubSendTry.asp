<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCode="GW1M51"
HTProgPrefix="ePub" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
Function Send_Email (S_Email,R_Email,Re_Sbj,Re_Body)
   Set objNewMail = CreateObject("CDONTS.NewMail") 
   objNewMail.MailFormat = 0
   objNewMail.BodyFormat = 0 
'   response.write S_Email&"[]"&R_Email&"[]"&Re_Sbj&"[]"&Re_Body
'   response.write "<hr>"
'   response.write Re_Body
'   response.end
   call objNewMail.Send(S_Email,R_Email,Re_Sbj,Re_Body)
   Set objNewMail = Nothing
End Function

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

Server.ScriptTimeOut = 2000

  on error resume next


 
 	epTreeID = session("epTreeID")		'-- 電子報的 tree

	ePubID = request.queryString("ePubID")
	formID = "ep" & ePubID
	
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
	oxsl.load(server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/ePaper"&epTreeID &".xsl"))
 
'response.write "<XMP>"+eXml.xml+"</XMP>"
'response.end
  	
  	'----參數處理
	ePaperTitle = dXml.selectSingleNode("ePaperTitle").text
	ePaperSEmail = dXml.selectSingleNode("ePaperSEmail").text
	ePaperSEmailName = dXml.selectSingleNode("ePaperSEmailName").text	
    	S_Email=""""&ePaperSEmailName&""" <"&ePaperSEmail&">"
	sCount = 0


	
	Response.ContentType = "text/HTML" 
	outString = replace(exml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	outString = replace(outString,"&amp;","&")
	xmBody = replace(outString,"&amp;","&")


        Email_body=xmBody



'---- 3. 送 只訂閱 ePaper 的 mail ---------------------------------------------------------------------
	sql = "SELECT u.* FROM ePaper2 AS u LEFT JOIN epSend AS s ON u.email=s.email" _
		& " AND s.ePubID=" & ePubID _
		& " WHERE s.email IS NULL"
	set RSlist = conn.execute(sql)
	
	while not RSlist.eof
            R_Email=RSlist("email")
            sCount = sCount + 1
            Call Send_Email(S_Email,R_Email,ePaperTitle,Email_body)
	
	    xsql = "INSERT INTO epSend(ePubID,email) VALUES(" _
		& ePubID & "," & pkStr(R_Email,")")
	    conn.execute xsql
	    RSlist.moveNext
	wend
%>
	<script language=vbs>	
		alert "電子報發送完成!  共發送<%=sCount%>份電子報"
		window.navigate "ePubEdit.asp?ePubID=<%=ePubID%>"	
	</script>		
	
