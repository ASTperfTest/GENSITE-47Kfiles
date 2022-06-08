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
	
	call lpSQLinjectionCheck()

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
	
	xv = oxml.load(session("myXDURL") & "/wsxd2/xdcp.asp?" & qStr)

  if oxml.parseError.reason <> "" then 
    Response.Write("htPageDom parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if
  
  xmyStyle=nullText(oxml.selectSingleNode("//hpMain/myStyle"))
  if xmyStyle="" then xmyStyle=session("myStyle")  
	oxsl.load(server.mappath("xslGip/" & xmyStyle & "/fp.xsl"))
	
  fStyle = oxml.selectSingleNode("//xslData").text
  if fStyle <> "" then
	set fxsl = server.createObject("microsoft.XMLDOM")
	fxsl.load(server.mappath("xslGip/" & fStyle & ".xsl"))
	'response.write "/xslGip/" & fStyle & ".xsl"
	set oxRoot = oxsl.selectSingleNode("xsl:stylesheet")
	
	on error resume next
	for each xs in fxsl.selectNodes("//xsl:template")
		set nx = xs.cloneNode(true)
		'response.write "aaa"
			ckStr = "@match='" & nx.getAttribute("match") & "'"
			if nx.getAttribute("mode")<>"" then		ckStr = ckStr & " and @mode='" & nx.getAttribute("mode") & "'"
			set orgEx = oxRoot.selectSingleNode("//xsl:template[" & ckStr & "]")
			oxRoot.removechild orgEx
			oxRoot.appendChild nx
		next
	end if  

	Response.ContentType = "text/HTML" 
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	outString = replace(outString," xmlns=""""","")
	response.write replace(outString,"&amp;","&")

	function nullText(xNode)
  	on error resume next
  	xstr = ""
  	xstr = xNode.text
  	nullText = xStr
	end function
%>
