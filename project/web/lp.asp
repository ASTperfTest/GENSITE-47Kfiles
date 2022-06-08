<%@ CodePage = 65001 %>
<%
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1
%>

<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#include file = "inc/web.sqlInjection.inc" -->
 <!--#include virtual = "/inc/checkURL.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
	next
	
	call lpSQLinjectionCheck()

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
			
	memID = session("memID")
	gstyle = session("gstyle")
	
	'response.write(session("myXDURL") & "/wsxd2/xdlp.asp?" & qStr & "&memID=" & memID & "&gstyle=" & gstyle)
	xv = oxml.load(session("myXDURL") & "/wsxd2/xdlp.asp?" & qStr & "&memID=" & memID & "&gstyle=" & gstyle)

    'Added By Leo  2011-07-13
    session("lpPageURL") = "/lp.asp?" & qStr
    if Request.QueryString("xq_xCat") <> "" then
        session("topCat") = Request.QueryString("xq_xCat")
    else
        session("topCat") = ""
    end if    
    'Added By Leo  2011-07-13
	'發生錯誤時，自動重整3次=====================================================
    %>
    <!--#include virtual = "/inc/OnErrorReload3Times.inc" -->
    <%
  '=============================================================================  

	xmyStyle = nullText(oxml.selectSingleNode("//hpMain/myStyle"))
  if xmyStyle = "" then xmyStyle = session("myStyle")  
	oxsl.load(server.mappath("xslGip/" & xmyStyle & "/lp.xsl"))
'response.Write server.mappath("xslGip/" & xmyStyle & "/lp.xsl")
  fStyle = oxml.selectSingleNode("//xslData").text
  if fStyle <> "" then
		set fxsl = server.createObject("microsoft.XMLDOM")
		fxsl.load(server.mappath("xslGip/" & fStyle & ".xsl"))		
		set oxRoot = oxsl.selectSingleNode("xsl:stylesheet")
  '  response.Write "<BR>"
  'response.Write server.mappath("xslGip/" & fStyle & ".xsl")
		on error resume next
		for each xs in fxsl.selectNodes("//xsl:template")
			set nx = xs.cloneNode(true)
			ckStr = "@match='" & nx.getAttribute("match") & "'"
			if nx.getAttribute("mode")<>"" then		ckStr = ckStr & " and @mode='" & nx.getAttribute("mode") & "'"
			set orgEx = oxRoot.selectSingleNode("//xsl:template[" & ckStr & "]")
			oxRoot.removechild orgEx
			oxRoot.appendChild nx
		next
		for each xs in fxsl.selectNodes("//msxsl:script")
			set nx = xs.cloneNode(true)
			oxRoot.appendChild nx
		next
	end if  

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

	response.write (outString)

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