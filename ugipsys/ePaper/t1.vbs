	DBConnStr="Provider=SQLOLEDB;server=10.10.5.63;User ID=hyGIP;Password=hyweb;Database=GIPmofDB"
	wadminRootPath = "c:\webSites\sMof"
	wadminRootPath = "d:\xWeb\sMof"


Function Send_Email (S_Email,R_Email,Re_Sbj,Re_Body)
   Set objNewMail = CreateObject("CDONTS.NewMail") 
   objNewMail.MailFormat = 0
   objNewMail.BodyFormat = 0 
'	wScript.Echo(R_Email)
   call objNewMail.Send(S_Email,R_Email,Re_Sbj,Re_Body)
   Set objNewMail = Nothing
End Function

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

FUNCTION pkStr (s, endchar)
  if s="" then
	pkStr = "null" & endchar
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	pkStr="'" & s & "'" & endchar
  end if
END FUNCTION


'Server.ScriptTimeOut = 2000

  on error resume next
 
 	epTreeID = 10		'-- 電子報的 tree
    S_Email="""財政部"" <internet@mail.mof.gov.tw>"
    xmBody = "haha"
    R_Email = "cwchen@hyweb.com.tw"
'	Call Send_Email(S_Email,R_Email,"財政部全球資訊網電子報xx",xmBody)

'wScript.Echo("open db conn..."&DBConnStr)
	Set Conn = CreateObject("ADODB.Connection")
	Conn.Open(DBConnStr)
'wScript.Echo("db conn opened.")
	
	Set RS = conn.execute("SELECT * FROM EpPub ORDER BY pubDate DESC")
	ePubID = RS("ePubID")
	formID = "ep" & ePubID
'wScript.Echo(formID)

	set oxml = createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty "ServerHTTPRequest", true
	LoadXML = wadminRootPath & "/public/ePaper/" & formID & ".xml"
	xv = oxml.load(LoadXML)
  if oxml.parseError.reason <> "" then
    wScript.Echo("XML parseError on line " &  oxml.parseError.line)
    wScript.Echo("<BR>Reason: " &  oxml.parseError.reason)
    WScript.Quit(1)
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
            wScript.Echo sCount & ") " & R_Email & "<BR>"
			xmBody = xmBody & dXml.selectSingleNode("footer").xml
            Call Send_Email(S_Email,R_Email,"財政部全球資訊網電子報",xmBody)
	
			xsql = "INSERT INTO epSend(ePubID,email) VALUES(" _
				& ePubID & "," & pkStr(R_Email,")")
			conn.execute xsql
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
            wScript.Echo sCount & ") " & R_Email & "<BR>"
			xmBody = xmBody & dXml.selectSingleNode("footer").xml
            Call Send_Email(S_Email,R_Email,"財政部全球資訊網電子報",xmBody)
	
			xsql = "INSERT INTO epSend(ePubID,email) VALUES(" _
				& ePubID & "," & pkStr(R_Email,")")
			conn.execute xsql
		  end if
'	response.end


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

