<%@ CodePage = 65001 %>
<%
	'Response.CacheControl = "no-cache" 
	'Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1
%>

<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#include file = "inc/web.sqlInjection.inc" -->
<!--#include virtual = "/inc/checkURL.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--Modified By Leo     2011-09-27  For 邀請朋友    Start    -->
<html>
    <head>
        <meta name="description" content="農業知識入口網讓您輕鬆習得農業知識，掌握最新農業資訊。加入農業知識入口網會員或登入會員，即有機會獲得大獎！還在等什麼呢？快點把大獎抱回家吧~" />
        <link rel="image_src" type="image/jpeg" href="images/FBIcon.png" />
    </head>
</html>
<!--Modified By Leo     2011-09-27  For 邀請朋友     End     -->
<?xml version="1.0"  encoding="utf-8" ?>
<%
	call CheckURL(Request.QueryString)
	mp = getMPvalue() 
	
	qStr = request.queryString
	if instr(qStr, "mp=") = 0 then qStr = qStr & "&mp=" & mp

	for each xf in request.form
		if pkStrWithSriptHTML( request(xf), "" ) <> "null" then 
			qStr = qStr & "&" & xf & "=" & server.URLEncode( stripHTML( request.form(xf) ) )
		end if
	Next
	
	If Not IsNull(session("CheckCodeforMail")) then
		qStr = qStr & "&" & "CheckCodeforMail" & "=" & server.URLEncode( session("CheckCodeforMail") )
	End if
	
	if left(qStr,1) = "&" Then qStr = mid(qStr,2)
	
	call spSQLinjectionCheck()

	memID = session("memID")
	gstyle = session("gstyle")

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
	

    'Response.Write(session("myXDURL") & "/wsxd2/xdSP.asp?" & qStr & "&memID=" & memID & "&gstyle=" & gstyle)
	'response.end
	xv = oxml.load(session("myXDURL") & "/wsxd2/xdSP.asp?" & qStr & "&memID=" & memID & "&gstyle=" & gstyle)


  if oxml.parseError.reason <> "" then 
    Response.Write("xhtPageDom parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if

  xmyStyle=nullText(oxml.selectSingleNode("//hpMain/myStyle"))
 	if xmyStyle="" then xmyStyle=session("myStyle")  
	oxsl.load(server.mappath("xslGip/" & xmyStyle & "/sp.xsl"))	

	Response.ContentType = "text/HTML" 
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">","")
	outString = replace(outString," xmlns=""""","")
	outString = replace(outString,"&amp;","&")
		  
	Dim memID, ShowCursorIcon
	ShowCursorIcon = "1"
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open session("ODBCDSN")
	memID = nullText(oxml.selectSingleNode("//hpMain/login/memID"))
	If (memID <> "") Then
			
			sql = "select ShowCursorIcon from Member Where account = '" & memID & "'"		
			Set loginrs = conn.execute(sql)
			If Not loginrs.Eof Then
				If Not IsNull(loginrs("ShowCursorIcon")) Then
					ShowCursorIcon = loginrs("ShowCursorIcon")
				else
					ShowCursorIcon = ChecCursorOpen
				End If
			End If
	else 
			ShowCursorIcon = ChecCursorOpen
	End If
	If ShowCursorIcon = "0" Then
				outString = replace(outString,"png.length!=0","false")
	End If
	
	response.write outString

	function nullText(xNode)
  	on error resume next
  	xstr = ""
  	xstr = xNode.text
  	nullText = xStr
	end function
	
	function ChecCursorOpen()
		sql = " select stitle from CuDTGeneric where icuitem = " & Application("ShowCursorIconId")
		Set RS = conn.execute(sql)
		If (Not IsNull(RS("sTitle")) ) and RS("sTitle") = 1 Then
			ChecCursorOpen = "1"
		else
			ChecCursorOpen = "0"
		End If
	end function
%>
