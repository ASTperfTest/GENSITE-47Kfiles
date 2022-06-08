<%@ CodePage = 65001 %><%
response.charset = "utf-8"
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1%>
<!--#include virtual = "/inc/checkURL.inc" -->
<!--#include virtual = "/subject/include/CheckPoint_BeforeLoadXML.asp" -->
<?xml version="1.0"  encoding="utf-8" ?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
call CheckURL(Request.QueryString)
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
	set oxsl = server.createObject("microsoft.XMLDOM")
	qStr = request.queryString
	if instr(qStr,"mp=")=0 then qStr = qStr & "&mp=" & session("mpTree")
	
	memID = session("memID")
	gstyle = session("gstyle")
	

'this function in include file
call CheckPoint_BeforeLoadXML(request("mp"), request("ctNode"), request("xItem"))
	
	
	xv = oxml.load(session("myXDURL") & "/wsxd/xdqp.asp?" & qStr & "&memID=" & memID & "&gstyle=" & gstyle)
	'response.write (session("myXDURL") & "/wsxd/xdqp.asp?" &qStr)
  if oxml.parseError.reason <> "" then 
    Response.Write("xhtPageDom parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if
    xmyStyle=nullText(oxml.selectSingleNode("//hpMain/MpStyle"))
	if xmyStyle="" then
        xmyStyle=nullText(oxml.selectSingleNode("//hpMain/myStyle"))
  		if xmyStyle="" then xmyStyle=session("myStyle") 
	end if 
	oxsl.load(server.mappath("xslGip/" & xmyStyle & "/qp.xsl"))
	

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
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">","")
	outString = replace(outString,"2007 All Rights Reserved",year(date) & " All Rights Reserved")
	response.write replace(outString,"&amp;","&")

'	response.write oxml.transformNode(oxsl)
'	response.write oxml.xml & "<HR>"
'	response.write oxsl.xml & "<HR>"
%>
