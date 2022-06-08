<%@ Language=VBScript %>
<% Response.Expires = 0 %>
<%
	formFunction = "query"

	formID = request("formID")
	progPath = request("progPath")
	if progPath <> "" then
		if left(progPath,1) = "/" then progPath = mid(progPath,2)
		progPath = replace(progPath,"/","\")
		progPath = progPath & "\"
	end if

	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true	
	
	LoadXML = server.MapPath(".") & "\formSpec\" & progPath & formID & ".xml"
	xv = htPageDom.load(LoadXML)
  if htPageDom.parseError.reason <> "" then 
    Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    Response.End()
  end if

pgPrefix = nullText(htPageDom.selectSingleNode("//htPage/HTProgPrefix"))
progPath = nullText(htPageDom.selectSingleNode("//htPage/HTProgPath"))
if progPath = "" then 
	pgPath = server.mapPath("genedCode/")
else
	pgPath = server.mapPath(progPath)
end if
If right(pgPath,1) <> "\" then pgPath = pgPath & "\"

	Dim xSearchListItem(20,2)
	xItemCount = 0

	set htFormDom = htPageDom.selectSingleNode("//pageSpec/formUI")
	set refModel = htPageDom.selectSingleNode("//htForm/formModel[@id='" & htFormDom.selectSingleNode("@ref").text & "']")

	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

	HTProgCap= nullText(htPageDom.selectSingleNode("//pageSpec/pageHead"))
	HTProgFunc=nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction"))
	HTProgCode=nullText(htPageDom.selectSingleNode("//htPage/HTProgCode"))
	HTProgPrefix=pgPrefix
	HTUploadPath="/"

 apath=server.mappath(HTUploadPath) & "\"
  set xup = Server.CreateObject("UpDownExpress.FileUpload")
  xup.Open 
    Set fs = CreateObject("scripting.filesystemobject")

function xUpForm(xvar)
on error resume next
	xStr = ""
	arrVal = xup.MultiVal(xvar)
	for i = 1 to Ubound(arrVal)
		xStr = xStr & arrVal(i) & ", "
'		Response.Write arrVal(i) & "<br>" & Chr(13)
	next 
	if xStr = "" then
		xStr = xup(xvar)
		xUpForm = xStr
	else
		xUpForm = left(xStr, len(xStr)-2)
	end if
end function

%>
<!--#Include virtual = "/inc/server.inc" -->
<!--#INCLUDE virtual="/inc/dbutil.inc" -->
<!--#Include file = "htUiGen.inc" -->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" href="/inc/setstyle.css">
</HEAD>
<BODY>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right>
<%
	for each x in htPageDom.selectNodes("//pageSpec/aidLinkList/Anchor")
		ckRight = nullText(x.selectSingleNode("checkRight"))
		if isNumeric(ckRight) then
			ckRight = cint(ckRight)
		else
			ckRight = 0
		end if
		if ckRight = 0 OR (HTProgRight and ckRight)= ckRight then
		  if x.selectSingleNode("AnchorType").text="Back" then
			response.write "<A href=""Javascript:window.history.back();"">" _
				& x.selectSingleNode("AnchorLabel").text & "</A> "
		  else
			anChorURI = x.selectSingleNode("AnchorURI").text
			xpos = inStrRev(anChorURI,".")
			if xpos>0 then	anChorURI = left(anChorURI,xpos-1)
			xPos = inStrRev(anChorURI,"/")
			if xPos>0 then
				xprogPath = left(anChorURI,xpos-1)
				anChorURI = mid(anChorURI,xPos+1)
			else
				xprogPath = request("progPath")
			end if
			response.write "<A href=""uiView.asp?formID=" & anChorURI & "&progPath=" & xprogPath _
				& """ title=""" & nullText(x.selectSingleNode("AnchorDesc")) & """>" _
				& x.selectSingleNode("AnchorLabel").text & "</A> "
		  end if
		end if
	next
%>		
    </td>    
    </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=90% height=230 valign=top>        

<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="">
<INPUT TYPE=hidden name=submitTask value="">

<%

   	calendarFlag=false
	for each x in htFormDom.selectSingleNode("pxHTML").childNodes
		recursiveQParam x
	next
   	if calendarFlag then
   		response.write "<object data=""/inc/calendar.htm"" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>"
		response.write "<INPUT TYPE=hidden name=CalendarTarget>"
   	end if

	for each xCode in htFormDom.selectNodes("scriptCode")
		response.write replace(xCode.text,chr(10),chr(13)&chr(10))
	next
	
    	if calendarFlag then
    		Set xfin = fs.opentextfile(Server.MapPath("Template0/TempCalendarSub.asp"))
			dumpTempFile    		
    		xfin.Close
    	end if
%>    	
</form>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
        <%if (HTProgRight AND 2) <> 0 then %>
            <input type=button value ="查　詢" class="cbutton" onClick="formSearchSubmit()">
            <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <%end if%>    
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">
 </td></tr>
</table>

</body>
</html>
