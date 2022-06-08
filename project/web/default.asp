<%@ CodePage = 65001 %>
<%
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1
%>

<!--#include file = "inc/web.sqlInjection.inc" -->
 <!--#include virtual = "/inc/checkURL.inc" -->
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
  	Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
  	Response.End()
  end if

  xs = nullText(htPageDom.selectSingleNode("//MpDataSet/MpStyle"))
	
	session("myStyle") = xs

	response.redirect "mp.asp?mp=1"

	function nullText(xNode)
	  on error resume next
	  xstr = ""
	  xstr = xNode.text
	  nullText = xStr
	end function

%>
