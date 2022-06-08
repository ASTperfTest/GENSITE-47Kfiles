<%@ CodePage = 65001 %><%
'Response.CacheControl = "no-cache" 
'Response.AddHeader "Pragma", "no-cache" 
Response.Expires = -1%>
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	
	qStr = request.queryString
	if instr(qStr,"mp=")=0 then qStr = qStr & "&mp=" & session("mpTree")
'	response.write qStr & "<HR>"
	for each xf in request.form
		if request(xf)<>"" then		qStr = qStr & "&" & xf & "=" & server.URLEncode(request.form(xf))
	next
	if left(qStr,1)="&"	then	qStr = mid(qStr,2)
'	response.write session("myXDURL") & "/wsxd2/xdOPML.asp?mp=" & session("mpTree") & "&" & qStr
'	response.end
	xv = oxml.load(session("myXDURL") & "/wsxd2/xdOPML.asp?" & qStr)
    if oxml.parseError.reason <> "" then 
        Response.Write("xhtPageDom parseError on line " &  oxml.parseError.line)
        Response.Write("<BR>Reason: " &  oxml.parseError.reason)
        Response.End()
    end if
        xmyStyle=nullText(oxml.selectSingleNode("//hpMain/myStyle"))
  	if xmyStyle="" then xmyStyle=session("myStyle")  

	Response.ContentType = "text/xml" 
	response.write replace(oxml.xml,"<?xml version=""1.0""?>","<?xml version=""1.0"" encoding=""UTF-8""?> ")
%>
