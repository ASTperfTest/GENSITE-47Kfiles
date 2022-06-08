<%@ Language=VBScript %>
<% Response.Expires = 0 %>
<%
	formFunction = "list"

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

    Set fs = CreateObject("scripting.filesystemobject")
	set htFormDom = htPageDom.selectSingleNode("//pageSpec")
	set refModel = htPageDom.selectSingleNode("//htPage/resultSet")
	set xDetail = htFormDom.selectSingleNode("detailRow")
'	response.write refModel.selectSingleNode("tableName").text & "<HR>"

	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

	HTProgCap= nullText(htPageDom.selectSingleNode("//pageSpec/pageHead"))
	HTProgFunc=nullText(htPageDom.selectSingleNode("//pageSpec/pageFunction"))
	HTProgCode=nullText(htPageDom.selectSingleNode("//htPage/HTProgCode"))
	HTProgPrefix=pgPrefix
%>
<!--#Include virtual = "/inc/server.inc" -->
<!--#INCLUDE virtual="/inc/dbutil.inc" -->
<!--#Include file = "htuiGen.inc" -->
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
	showAidLinkList
%>		
    </td>    
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
 <Form name=reg method="POST" action=AssociationList.asp>
  <p align="center">  
     
     <font size="2" color="rgb(63,142,186)"> 第
     <font size="2" color="#FF0000">1/1</font>頁|                      
        共<font size="2" color="#FF0000">2</font>
       <font size="2" color="rgb(63,142,186)">筆| 跳至第       
         <select id=GoPage size="1" style="color:#FF0000">
              
                   <option value="1"selected>1</option>          
                
         </select>      
         頁</font>           
            
        | 每頁筆數:
       <select id=PerPage size="1" style="color:#FF0000">            
             <option value="10" selected>10</option>                       
             <option value="20">20</option>
             <option value="30">30</option>
             <option value="50">50</option>
        </select>     
     </font>     
    <CENTER>
     <TABLE width=100% cellspacing="1" cellpadding="0" class=bg>                   
     <tr align=left>    
<%
	for each param in xDetail.selectNodes("colSpec")
		response.write "<td class=eTableLable>" & nullText(param.selectSingleNode("colLabel")) & "</td>"
	next
	response.write "</tr>"

	xpLineStr = "<tr>"
	for each param in xDetail.selectNodes("colSpec")
		xpLineStr = xpLineStr & "<TD class=eTableContent><font size=2>" 
		xUrl = nullText(param.selectSingleNode("url"))
		if  xUrl <> "" then
			anChorURI = xUrl
			xpos = inStrRev(anChorURI,".")
			if xpos>0 then	anChorURI = left(anChorURI,xpos-1)
			xPos = inStrRev(anChorURI,"/")
			if xPos>0 then
				xprogPath = left(anChorURI,xpos-1)
				anChorURI = mid(anChorURI,xPos+1)
			else
				xprogPath = request("progPath")
			end if
			xpLineStr = xpLineStr & "<A href=""uiView.asp?formID=" & anChorURI & "&progPath=" & xprogPath & """>"
		end if
		xpLineStr = xpLineStr & "xxxxxxxx"
		if xUrl <> "" then	xpLineStr = xpLineStr & "</A>"
		xpLineStr = xpLineStr & "</font></td>"
	next

	for xi=1 to 3    
		response.write xpLineStr
	next

%>
</table>
<% showUINotes %>
</BODY>
</HTML>
