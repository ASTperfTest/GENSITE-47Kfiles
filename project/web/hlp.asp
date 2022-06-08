<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--
<div class="top">
			<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0"
				width="780" height="150" VIEWASTEXT ID="Object1">
				<param name="movie" value="css/images/innerhead.swf">
				<param name="quality" value="high">
				<embed src="css/images/innerhead.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer"
					type="application/x-shockwave-flash" width="780" height="150"></embed>
				
			</object>
		</div>
-->
<%
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1%><?xml version="1.0"  encoding="big5" ?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!--#include virtual = "/inc/checkURL.inc" -->
<!--#include file = "inc/web.sqlInjection.inc" -->
<%
	on error resume next
	call CheckURL(Request.QueryString)
'-- HitCount start here ------------------------------------------------
	xCtUnit = cint(request("CtUnit"))
	if xCtUnit <> session("gHitUnitID") then
	  Set Conn = Server.CreateObject("ADODB.Connection")
	  Conn.Open session("ODBCDSN")
	  if session("gHitUnitID") >0 then
	  	staySec = datediff("s", session("gHitUnitTime"), now())
		sql = "UPDATE gipHitUnit SET staySec = " & staySec _
			& " WHERE gHitSID=" & session("gHitSID") _
			& " AND iCtUnit=" & session("gHitUnitID") _
			& " AND staySec IS NULL"
		conn.execute sql
	  end if
		session("gHitUnitID") = xCtUnit
		session("gHitUnitTime") = now()
		xNode = request("ctNode")
		if xNode="" then	xNode = session("gHitNode")
		if xNode="" then	xNode = 0
		sql = "INSERT INTO gipHitUnit(gHitSID, iCtNode, iCtUnit) VALUES(" _
			& session("gHitSID") & "," & xNode & "," & xCtUnit & ")"
		conn.execute sql
		session("gHitLast") = now()
		session("gHitVcUnit") = session("gHitVcUnit") + 1
	end if
	if request("ctNode")<>"" then	session("gHitNode") = request("ctNode")
'-- HitCount done here ------------------------------------------------

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
	mp = getMPvalue()
		if instr(qStr, "mp=") = 0 then qStr ="mp=" & mp
	qStr = qStr & "&" & request.queryString
	for each xf in request.form
		if request(xf)<>"" then
                                     Session("formValue")=request.form(xf)
                                     Session.CodePage=950		
qStr = qStr & "&" & xf & "=" & server.URLEncode(Session("formValue"))
end if
	next
                   Session.CodePage=65001
	'qStr = server.URLEncode(qStr)
	 style=request.cookies("style1") 
	xv = oxml.load(session("myXdURL") & "/wsxd2/xdhlp.asp?"& qStr)
	'response.write session("myXdURL") & "/wsxd2/xdhlp.asp?" & qStr 
	'response.write oxml.xml
	'response.end
  if oxml.parseError.reason <> "" then 
  	response.write "<html>"
	response.write session("myXDURL") & "/wsxd2/xdhlp.asp?"& qStr
    Response.Write("<BR/>htPageDom parseError on line " &  oxml.parseError.line)
    Response.Write("<BR/>Reason: " &  oxml.parseError.reason)
  	response.write "</html>"
    Response.End()
  end if
	oxsl.load(server.mappath("xslGip/" & session("myStyle") & "/2sidepage.xsl"))
	'response.write "xslGip/" & session("myStyle") & "/2sidepage.xsl"
  	fStyle = oxml.selectSingleNode("//xslData").text
  	if fStyle <> "" then
		set fxsl = server.createObject("microsoft.XMLDOM")
		fxsl.load(server.mappath("xslGip/" & fStyle & ".xsl"))
		set oxRoot = oxsl.selectSingleNode("xsl:stylesheet")
	
	on error resume next
		for each xs in fxsl.selectNodes("//xsl:template")
			set nx = xs.cloneNode(true)
			ckStr = "@match='" & nx.getAttribute("match") & "'"
			if nx.getAttribute("mode")<>"" then		ckStr = ckStr & " and @mode='" & nx.getAttribute("mode") & "'"
			set orgEx = oxRoot.selectSingleNode("//xsl:template[" & ckStr & "]")
			oxRoot.removechild orgEx
			oxRoot.appendChild nx
		next
	end if  

'	response.write oxml.xml
'	response.end
	Response.ContentType = "text/HTML" 
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	
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
	
	function ChecCursorOpen()
		sql = " select stitle from CuDTGeneric where icuitem = " & Application("ShowCursorIconId")
		Set RS = conn.execute(sql)
		If (Not IsNull(RS("sTitle")) ) and RS("sTitle") = 1 Then
			ChecCursorOpen = "1"
		else
			ChecCursorOpen = "0"
		End If
	end function

'	response.write oxml.transformNode(oxsl)
'	response.write oxml.xml & "<HR>"
'	response.write oxsl.xml & "<HR>"
%>
