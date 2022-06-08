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
<!--#Include file = "htUIBGen.inc" -->
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
 <Form name=reg method="POST">
    <td class="FormLink" valign="top" align=left>
<%
	showAidLinkList
%>		
    </td>    
	<td  id=SJob class="Formtext" width="50%" align=right>
	   <SPAN id=RunJob style="visibility:hidden">
	   	   <input type=button class=cbutton value="確定執行" id=button1 name=button1>
	   	   <INPUT TYPE=HIDDEN name=doJob VALUE="">
	   </SPAN>
	</td>	    
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
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
	<td align=center class=eTableLable width=7%>
	<input type=button value ="全選" class="cbutton"  name=ckall onClick="ChkAll"></td>
<%
	for each param in xDetail.selectNodes("colSpec")
		response.write "<td class=eTableLable>" & nullText(param.selectSingleNode("colLabel")) & "</td>"
	next
	response.write "</tr>"

	xpLineStr = ""
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
		processContent param.selectSingleNode("content")
'		if nullText(param.selectSingleNode("url"))
'		xpLineStr = xpLineStr & "xxxxxxxx"
		if xUrl <> "" then	xpLineStr = xpLineStr & "</A>"
		xpLineStr = xpLineStr & "</font></td>"
	next

	for xi=1 to 3    
		response.write "<tr><TD class=eTableContent align=center>" _
			& "<input type=checkbox name=ckbox" & xi & "></td>"
		response.write xpLineStr & "</tr>"
	next

%>
</table>
<% showUINotes %>
</BODY>
<script language=vbs>   
      Dim chkCount
      chkCount=0            '記錄checkbox 被勾數
   
      document.all.SelectJob.value="<%=fOpt%>"

    sub document_onClick           'checkbox 被勾計數
         set sObj=window.event.srcElement
         if sObj.tagName="INPUT" then 
            if sObj.type="checkbox"  then 
                if sObj.checked then 
                   chkCount=chkCount+1
                else
                   chkCount=chkCount-1                
                end if                                          
            end if
         end if
         '
         if chkCount=0 then 
            document.all("RunJob").style.visibility="hidden"
         else
            document.all("RunJob").style.visibility="visible"
         end if
    end sub        
    
     
    sub SelectJob_onChange           '作業狀態改變時處理

		anChorURI = document.all("SelectJob").value
		xpos = inStrRev(anChorURI,".")
		if xpos>0 then	anChorURI = left(anChorURI,xpos-1)
		xPos = inStrRev(anChorURI,"/")
		if xPos>0 then
			xprogPath = left(anChorURI,xpos-1)
			anChorURI = mid(anChorURI,xPos+1)
		else
			xprogPath = "<%=request("progPath")%>"
		end if

		document.location.href= "uiView.asp?formID=" & anChorURI & "&progPath=" & xprogPath
    end sub   


   sub RunJob_Onclick                   ''確定執行 打勾項              
       reg.doJob.value = "<%=fOpt%>"
       reg.submit
   end sub  
  

   sub GoPage_OnChange               '跳至頁數
        newPage=reg.GoPage.value     
        document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&pagesize=<%=PerPageSize%>&opt=<%=fOpt%>"    
   end sub      
     
   sub PerPage_OnChange              '指定每頁筆數  
        newPerPage=reg.PerPage.value
        document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>" & "&pagesize=" & newPerPage & "&opt=<%=fOpt%>"
   end sub 
   
   sub Chkall
       chkCount=0     
       if document.all("ckall").value="全選" then           '全勾
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=true
                   chkCount=chkCount+1
              end if     
          next                 
          document.all("RunJob").style.visibility="visible"
          document.all("ckall").value="全不選"
      elseif document.all("ckall").value="全不選" then        '全不勾
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=false
               end if     
          next                 
          document.all("RunJob").style.visibility="hidden"
          document.all("ckall").value="全選"          
       end if
   end sub   
</script>

</HTML>
<%
sub processContent(xDom)
dim x
	if xDom.nodeName = "refField" then
		processRefField xDom.text
'		xpLineStr = xpLineStr & "xxxxxxxx"
'	  	xfout.writeLine cl & "=RSreg(""" & xDom.text & """)" & cr
		exit sub
	end if
	if xDom.nodeName = "#comment" then	exit sub
	if xDom.nodeName = "#text" then
		xpLineStr = xpLineStr & xDom.text
'		xfout.writeLine xDom.text
		exit sub
	end if
	for each x in xDom.childNodes
		processContent x
	next
end sub

sub processRefField (xRefField)
	set xrf = refModel.selectSingleNode("fieldList/field[fieldName='" & xRefField & "']")
	if nullText(xrf.selectSingleNode("valueType")) = "direct" then
		xpLineStr = xpLineStr & "xxxxxxxx"
	else
		processParamField xrf
	end if
end sub
%>