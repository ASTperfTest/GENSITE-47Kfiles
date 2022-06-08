<%@ CodePage = 65001 %>
<%
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1
%>

<!--#include file = "inc/web.sqlInjection.inc" -->
<!--#include virtual = "/inc/checkURL.inc" -->

<?xml version="1.0"  encoding="utf-8" ?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
call CheckURL(Request.QueryString)
	mp = getMPvalue() 

	call mpSQLinjectionCheck()

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
	
	xv = oxml.load(session("myXDURL") & "/wsxd2/xdsitemap.asp?mp=" & mp)
	
  if oxml.parseError.reason <> "" then 
    Response.Write("htPageDom parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if
  
 	xmyStyle=nullText(oxml.selectSingleNode("//hpMain/myStyle"))
  if xmyStyle="" then xmyStyle=session("myStyle") 
	oxsl.load(server.mappath("xslGip/" & xmyStyle & "/sitemap.xsl"))
	
	Response.ContentType = "text/HTML" 
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">","")
	outString = replace(outString,"&amp;","&")
	response.write replace(outString,"&amp;","&")
	
	
	function nullText(xNode)
  	on error resume next
  	xstr = ""
  	xstr = xNode.text
  	nullText = xStr
	end function
	
%>
