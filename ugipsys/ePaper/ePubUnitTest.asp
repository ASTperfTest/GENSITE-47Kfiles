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
'   call objNewMail.Send(S_Email,"mjchen@mail.mof.gov.tw",Re_Sbj,Re_Body&" sendTo:- "&R_Email)
   call objNewMail.Send(S_Email,"cwchen@hyweb.com.tw",Re_Sbj,Re_Body&" sendTo:- "&R_Email)
	response.write R_Email & "<HR>"
   Set objNewMail = Nothing
End Function

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

  on error resume next
 
 	epTreeID = 10		'-- 電子報的 tree

	ePubID = request.queryString("ePubID")
	formID = "ep" & ePubID

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/public/ePaper/" & formID & ".xml")
	xv = oxml.load(LoadXML)
  if oxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if
  
  	set dXml = oxml.selectSingleNode("ePaper")
  	xmBody = ""
  	dXml.selectSingleNode("header//ePubDate").text = date()
	gxmBody = xmBody & dXml.selectSingleNode("header").xml

    S_Email="""財政部"" <internet@mail.mof.gov.tw>"

	xmBody = gxmBody
	sCount = 0

'---- 1. 送客製化選單元的 mail ---------------------------------------------------------------------
	sql = "SELECT u.*, p.ctNodeID FROM member AS u JOIN memEPaper AS p ON u.account=p.memID" _
		& " JOIN CatTreeNode AS n ON p.CtNodeID=n.CtNodeID" _
		& " LEFT JOIN epSend AS s ON u.email=s.email" & " AND s.ePubID=" & ePubID _
		& " WHERE s.email IS NULL" _
		& " ORDER BY u.account, n.CatShowOrder"
	set RSlist = conn.execute(sql)
'	response.write sql & "<HR>"

	xUser = ""
	while not RSlist.eof
'		response.write RSlist("account") & RSlist("ctNodeID") & "<BR>"
		if RSlist("account") <> xUser then
		  if xUser <> "" then
            sCount = sCount + 1
            response.write sCount & ") " & R_Email & "<BR>"
			xmBody = xmBody & dXml.selectSingleNode("footer").xml
            Call Send_Email(S_Email,R_Email,"財政部全球資訊網電子報",xmBody)
	
			xsql = "INSERT INTO epSend(ePubID,email) VALUES(" _
				& ePubID & "," & pkStr(R_Email,")")
'			conn.execute xsql
		  end if
		  	xUser = RSlist("account")
            R_Email=RSlist("email")
			xmBody = gxmBody
		end if
		
		xmBody=xmBody & nullText(dXml.selectSingleNode("epSectionList/epSection[secID='" & RSlist("ctNodeID") & "']/secBody"))
		
		RSlist.moveNext
	wend
		  if xUser <> "" then
            sCount = sCount + 1
            response.write sCount & ") " & R_Email & "<BR>"
			xmBody = xmBody & dXml.selectSingleNode("footer").xml
            Call Send_Email(S_Email,R_Email,"財政部全球資訊網電子報",xmBody)
	
			xsql = "INSERT INTO epSend(ePubID,email) VALUES(" _
				& ePubID & "," & pkStr(R_Email,")")
'			conn.execute xsql
		  end if
	response.end

	xmBody = gxmBody
	for each xSec in dXml.selectNodes("epSectionList/epSection")
'		response.write xSec.selectSingleNode("secBody").text
		xmBody = xmBody & xSec.selectSingleNode("secBody").text
	next
	xmBody = xmBody & dXml.selectSingleNode("footer").xml
'		response.write xmBody

            S_Email="""財政部"" <internet@mail.mof.gov.tw>"
'            R_Email=rs("Email")
'            R_Email="cwchen@hyweb.com.tw"
'            R_Email="mjchen@mail.mof.gov.tw"
           Email_body=xmBody

'---- 2. 送 Member 的 mail ---------------------------------------------------------------------
'	sCount = 0
	sql = "SELECT u.* FROM member AS u LEFT JOIN epSend AS s ON u.email=s.email" _
		& " AND s.ePubID=" & ePubID _
		& " WHERE s.email IS NULL"
	set RSlist = conn.execute(sql)
	
	while not RSlist.eof
            R_Email=RSlist("email")
            sCount = sCount + 1
            response.write sCount & ") " & R_Email & "<BR>"
            Call Send_Email(S_Email,R_Email,"財政部全球資訊網電子報",Email_body)
	
			xsql = "INSERT INTO epSend(ePubID,email) VALUES(" _
				& ePubID & "," & pkStr(R_Email,")")
			conn.execute xsql
		RSlist.moveNext
	wend

'---- 3. 送 只訂閱 ePaper 的 mail ---------------------------------------------------------------------
	sql = "SELECT u.* FROM ePaper AS u LEFT JOIN epSend AS s ON u.email=s.email" _
		& " AND s.ePubID=" & ePubID _
		& " WHERE s.email IS NULL"
	set RSlist = conn.execute(sql)
	
	while not RSlist.eof
            R_Email=RSlist("email")
            sCount = sCount + 1
            response.write sCount & ") " & R_Email & "<BR>"
            Call Send_Email(S_Email,R_Email,"財政部全球資訊網電子報",Email_body)
	
			xsql = "INSERT INTO epSend(ePubID,email) VALUES(" _
				& ePubID & "," & pkStr(R_Email,")")
			conn.execute xsql
		RSlist.moveNext
	wend


	
	
	
	response.end



%>