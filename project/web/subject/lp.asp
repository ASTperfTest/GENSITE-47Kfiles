<%@ CodePage = 65001 %>
<%
	response.charset = "utf-8"
    Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1
%>
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/web.sqlInjection.inc" -->
<!--#include virtual = "/inc/checkURL.inc" -->
<!--#include virtual = "/subject/include/CheckPoint_BeforeLoadXML.asp" -->
<%

htx_xpostDateS = request("htx_xpostDateS")
htx_xpostDateE = request("htx_xpostDateE")

if htx_xpostDateS <> "" or htx_xpostDateE <>"" then
    isDateValid = isdate(htx_xpostDateS)
    isDateValid = isdate(htx_xpostDateE) and isDateValid
    if not isDateValid then
        %>
        <html>
        <head><title></title></head>
        <body>
            <script type="text/javascript">
                alert('日期格式錯誤');
                window.location.href = "qp.asp?<%=request.ServerVariables("QUERY_STRING") %>"                
            </script>
        </body>
        </html>
        <%        
        response.End
    end if
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<?xml version="1.0"  encoding="utf-8" ?>
<%

call CheckURL(Request.QueryString)
	mp = getMPvalue() 

	qStr = request.queryString
	if instr(qStr, "mp=") = 0 then qStr = qStr & "&mp=" & mp
	
	for each xf in request.form	
		if pkStrWithSriptHTML( request(xf), "" ) <> "null" then 
		    value = server.URLEncode( stripHTML( request.form(xf) ) )		    
			qStr = qStr & "&" & xf & "=" & value
	        '判斷是不是由條件式查詢來的	        
		    if left(xf, 4) = "htx_" and value <> "" and left(qStr, 15) <> "isUserSearch=Y&" then
		        qStr = "isUserSearch=Y&" & qStr
		    end if
		end if
	next
	call lpSQLinjectionCheck()

'this function in include file
call CheckPoint_BeforeLoadXML(request("mp"), request("ctNode"), request("xItem"))

	
	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
	
	memID = session("memID")
	gstyle = session("gstyle")
	
	xv = oxml.load(session("myXDURL") & "/wsxd/xdlp.asp?" & qStr & "&memID=" & memID & "&gstyle=" & gstyle)

  if oxml.parseError.reason <> "" then 
  	response.write "<html>"
    Response.Write("<BR/>htPageDom parseError on line " &  oxml.parseError.line)
    Response.Write("<BR/>Reason: " &  oxml.parseError.reason)
  	response.write "</html>"
    'Response.End()
  end if
 	xmyStyle= nullText(oxml.selectSingleNode("//hpMain/MpStyle"))
	if xmyStyle="" then	
		xmyStyle=nullText(oxml.selectSingleNode("//hpMain/myStyle"))
 		if xmyStyle="" then xmyStyle=session("myStyle")
	end if
	
	oxsl.load(server.mappath("xslGip/" & xmyStyle & "/lp.xsl"))
	
  	fStyle = oxml.selectSingleNode("//xslData").text
  	if fStyle <> "" then
		set fxsl = server.createObject("microsoft.XMLDOM")
		fxsl.load(server.mappath("xslGip/" & xmyStyle & "/" & fStyle & ".xsl"))
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
		for each xs in fxsl.selectNodes("//msxsl:script")
			set nx = xs.cloneNode(true)
			oxRoot.appendChild nx
		next
	end if  

	Response.ContentType = "text/HTML" 
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">","")
	outString=replace(outString,"&amp;","&")
	outString=replace(outString,"&lt;","<")
	outString=replace(outString,"&gt;",">")	
	
	'增加計數用程式, Bike 2009/09/21
	outString = replace(outString,"</body>","<script type=""text/javascript"" src=""/ViewCounter.aspx""></script></body>")
	
	outString = replace(outString,"2007 All Rights Reserved",year(date) & " All Rights Reserved")
	response.write outString

	function nullText(xNode)
		on error resume next
		xstr = ""
		xstr = xNode.text
		nullText = xStr
	end function
%>
