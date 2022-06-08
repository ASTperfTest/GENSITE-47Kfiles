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
<%
	call CheckURL(Request.QueryString)
	mp = getMPvalue() 
	
	qStr = request.queryString
	if instr(qStr, "mp=") = 0 then qStr = qStr & "&mp=" & mp

	for each xf in request.form
		if pkStrWithSriptHTML( request(xf), "" ) <> "null" then 
			qStr = qStr & "&" & xf & "=" & server.URLEncode( stripHTML( request.form(xf) ) )
		end if
	next
	
	if left(qStr,1) = "&"	then	qStr = mid(qStr,2)
	
	call spSQLinjectionCheck()

	memID = session("memID")
	gstyle = session("gstyle")

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")	
	xv = oxml.load(session("myXDURL") & "/wsxd2/xdHSP.asp?" & qStr & "&memID=" & memID & "&gstyle=" & gstyle)
	 'Response.Write session("myXDURL") & "/wsxd2/xdHSP.asp?" & qStr & "&memID=" & memID & "&gstyle=" & gstyle


  if oxml.parseError.reason <> "" then 
    Response.Write("xhtPageDom parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if
 	
 	xmyStyle=nullText(oxml.selectSingleNode("//hpMain/myStyle"))
  if xmyStyle="" then xmyStyle=session("myStyle")  
	oxsl.load(server.mappath("xslGip/" & xmyStyle & "/2sidepage.xsl"))	

	Response.ContentType = "text/HTML" 
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">","")
	outString = replace(outString," xmlns=""""","")
	
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
	
	response.write replace(outString,"&amp;","&")

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
