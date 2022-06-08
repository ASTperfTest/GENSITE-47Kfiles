<%@ CodePage = 65001 %>
<%
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1
%>

<!--#include file = "inc/web.sqlInjection.inc" -->
 <!--#include virtual = "/inc/checkURL.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<?xml version="1.0"  encoding="utf-8" ?>
<%
call CheckURL(Request.QueryString)
	mp = getMPvalue() 

	call mpSQLinjectionCheck()

	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true	
	
	LoadXML = session("myXDURL") & "/GipDSD/xdmp" & mp & ".xml"
	xv = htPageDom.load(LoadXML)
		
  if htPageDom.parseError.reason <> "" then 
  	Response.Write("xdmp1.xml parseError on line " &  htPageDom.parseError.line)
  	Response.Write("<BR>Reason: " &  htPageDom.parseError.reason)
  	Response.End()
  end if

  xs = nullText(htPageDom.selectSingleNode("//MpDataSet/MpStyle"))
	
	session("myStyle") = xs

	response.redirect "mp.asp?mp="& session("mpTree")

	function nullText(xNode)
  	on error resume next
  	xstr = ""
  	xstr = xNode.text
  	nullText = xStr
	end function

%>
