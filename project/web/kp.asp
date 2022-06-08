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
	xv = oxml.load(session("myXDURL") & "/wsxd2/xdSP.asp?" & qStr & "&memID=" & memID & "&gstyle=" & gstyle)
	

  '發生錯誤時，自動重整3次=====================================================
    %>
    <!--#include virtual = "/inc/OnErrorReload3Times.inc" -->
    <%
  '=============================================================================

 	
 	xmyStyle=nullText(oxml.selectSingleNode("//hpMain/myStyle"))
  if xmyStyle="" then xmyStyle=session("myStyle")  
	oxsl.load(server.mappath("xslGip/" & xmyStyle & "/2sidepage.xsl"))	

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
	
	Function URLEncode(strEnc)   
     Dim strChr, intAsc, strTmp, strTmp2, strRet, lngLoop   
     For lngLoop = 1 To Len(strEnc)   
         strChr = Mid(strEnc, lngLoop, 1)   
         intAsc = Asc(strChr)   
         If ((intAsc < 58) And (intAsc > 47)) Or ((intAsc < 91) And _   
                (intAsc > 64)) Or ((intAsc < 123) And (intAsc > 96)) Then  
             strRet = strRet & strChr   
         ElseIf intAsc = 32 Then  
             strRet = strRet & "+"  
         Else  
             strTmp = Hex(Asc(strChr))   
             strRet = strRet & "%" & Right("00" & Left(strTmp, 2), 2)   
             strTmp2 = Mid(strTmp, 3, 2)   
             If Len(strTmp) > 3 Then  
                 If IsNumeric(Mid(strTmp, 3, 1)) Then  
                     strRet = strRet & Chr(CInt("&H" & strTmp2))   
                 Else  
                     strRet = strRet & "%" & strTmp2   
                 End If  
             End If  
         End If  
     Next  
     URLEncode = strRet   
 End Function
%>
