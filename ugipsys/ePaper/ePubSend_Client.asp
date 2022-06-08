<%@ CodePage = 65001 %>
<%
'// purpose: to send epaper without back side login.
'// modify date: 2006/01/06
'// ps: this asp include client.inc instead of server.inc
'// =========================================

Response.Charset="utf-8"
Response.Expires = 0
HTProgCode="GW1M51"
HTProgPrefix="ePub" 
%>
<!--#include virtual = "/inc/client.inc" -->
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
	
	'response.write server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/ePaper"&epTreeID &".xsl")
	'response.end
	
	oxsl.load(server.MapPath("/site/" & session("mySiteID") & "/public/ePaper/ePaper"&epTreeID &".xsl"))
 
'response.write "<XMP>"+eXml.xml+"</XMP>"

  	
  	'----參數處理
	ePaperTitle = dXml.selectSingleNode("ePaperTitle").text
	ePaperSEmail = dXml.selectSingleNode("ePaperSEmail").text
	ePaperSEmailName = dXml.selectSingleNode("ePaperSEmailName").text	
    	S_Email=""""&ePaperSEmailName&""" <"&ePaperSEmail&">"
	sCount = 0

'---- 1. 送客製化選單元的 mail ---------------------------------------------------------------------
	sql = "SELECT u.*, p.ctNodeId FROM Member AS u JOIN MemEpaper AS p ON u.account=p.memId" _
		& " JOIN CatTreeNode AS n ON p.ctNodeId=n.ctNodeId" _
		& " LEFT JOIN EpSend AS s ON u.email=s.email" & " AND s.epubId=" & ePubID & " and s.ctRootId=" & session("epTreeID") _
		& " WHERE s.email IS NULL" _
		& " ORDER BY u.account, n.catShowOrder"
'response.write sql
	
	set RSlist = conn.execute(sql)
	xUser = ""
	set eXml = oxml.cloneNode(true)	
	set epSectionListNode =	oxml.selectSingleNode("ePaper/epSectionList").cloneNode(true)	
	set nxml0 = server.createObject("microsoft.XMLDOM")
	nxml0.LoadXML("<epSectionList></epSectionList>")
	set newNode = nxml0.documentElement 
'response.write "<XMP>"+eXml.xml+"</XMP>"
'response.end	
	if not RSlist.eof then
	    while not RSlist.eof
		if RSlist("account") <> xUser then
		   if xUser <> "" then
            		sCount = sCount + 1
            		set ePaperNode = eXml.selectSingleNode("ePaper")
            		ePaperNode.removeChild ePaperNode.selectSingleNode("epSectionList")
            		ePaperNode.appendChild newNode
			Response.ContentType = "text/HTML" 
			outString = replace(exml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
			outString = replace(outString,"&amp;","&")
			xmBody = replace(outString,"&amp;","&")
                      
      
            		Call Send_Email(S_Email,R_Email,ePaperTitle,xmBody)
			set eXml = oxml.cloneNode(true)	
			set epSectionListNode =	oxml.selectSingleNode("ePaper/epSectionList").cloneNode(true)	
			set nxml0 = server.createObject("microsoft.XMLDOM")
			nxml0.LoadXML("<epSectionList></epSectionList>")
			set newNode = nxml0.documentElement
			xsql = "INSERT INTO epSend(ePubID,email) VALUES(" _
				& ePubID & "," & pkStr(R_Email,")")
			conn.execute xsql
		   end if
		   xUser = RSlist("account")
            	   R_Email=RSlist("email")
		end if
		
		newNode.appendChild epSectionListNode.selectSingleNode("epSection[secID='" & RSlist("ctNodeID") & "']")
		
		RSlist.moveNext
	    wend
	    if xUser <> "" then
            	sCount = sCount + 1
            		set ePaperNode = eXml.selectSingleNode("ePaper")
            		ePaperNode.removeChild ePaperNode.selectSingleNode("epSectionList")
            		ePaperNode.appendChild newNode
			Response.ContentType = "text/HTML" 
			outString = replace(exml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
			outString = replace(outString,"&amp;","&")
			xmBody = replace(outString,"&amp;","&")
           
            	Call Send_Email(S_Email,R_Email,ePaperTitle,xmBody)
	
	    	xsql = "INSERT INTO epSend(ePubID,email,CtRootID) VALUES(" _
				& ePubID & ",'" & R_Email & "',"&session("epTreeID")&")"
	    	conn.execute xsql
	    end if
	end if
	'-----客製化email發送完成
	
	Response.ContentType = "text/HTML" 
	outString = replace(exml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	outString = replace(outString,"&amp;","&")
	xmBody = replace(outString,"&amp;","&")


        Email_body=xmBody
'response.end


'---- 3. 送 只訂閱 ePaper 的 mail ---------------------------------------------------------------------
	sql = "SELECT u.* FROM Epaper AS u LEFT JOIN EpSend AS s ON u.email=s.email and u.ctRootId=s.ctRootId" _
		& " AND s.epubId=" & ePubID & " where u.ctRootId=" & session("epTreeID") _
		& " and s.email IS NULL"
'response.write sql
'response.end		
	set RSlist = conn.execute(sql)
	
	while not RSlist.eof
            R_Email=RSlist("email")
            sCount = sCount + 1
            Call Send_Email(S_Email,R_Email,ePaperTitle,Email_body)
	
	    xsql = "INSERT INTO epSend(ePubID,email,CtRootID) VALUES(" _
				& ePubID & ",'" & R_Email & "',"&session("epTreeID")&")"
	    conn.execute xsql
	    RSlist.moveNext
	wend
%>
	<script language=vbs>	
		alert "電子報發送完成!  共發送<%=sCount%>份電子報"
		window.navigate "ePubEdit.asp?ePubID=<%=ePubID%>&phase=edit"	
	</script>		
	
