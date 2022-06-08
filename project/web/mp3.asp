<%@ CodePage = 65001 %>
<%
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1
%>

<!--#include file = "inc/web.sqlInjection.inc" -->
 <!--#include virtual = "/inc/checkURL.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<?xml version="1.0"  encoding="utf-8" ?>
<%
call CheckURL(Request.QueryString)
	mp = getMPvalue() 

	call mpSQLinjectionCheck()
	
	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
	
	memID = session("memID")
	gstyle = session("gstyle")
	
	xv = oxml.load(session("myXDURL") & "/wsxd2/xdmp.asp?mp=" & mp & "&memID=" & memID & "&gstyle=" & gstyle)
		
  if oxml.parseError.reason <> "" then 
    Response.Write("htPageDom parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if
       
  xmyStyle = nullText(oxml.selectSingleNode("//hpMain/myStyle"))
  if xmyStyle = "" then xmyStyle = session("myStyle")
	oxsl.load(server.mappath("xslGip/" & xmyStyle & "/mp3.xsl"))
  if oxsl.parseError.reason <> "" then 
    Response.Write("oxslhtPageDom parseError on line " &  oxsl.parseError.line)
    Response.Write("<BR>Reason: " &  oxsl.parseError.reason)
    Response.End()
  end if
	
	Response.ContentType = "text/HTML" 

	'response.write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">" & vbCRLF
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">","")
	outString = replace(outString," xmlns=""""","")
	outString = replace(outString,"&amp;","&")
	response.write replace(outString,"&amp;","&")

	function nullText(xNode)
  	on error resume next
  	xstr = ""
  	xstr = xNode.text
  	nullText = xStr
	end function

%>
