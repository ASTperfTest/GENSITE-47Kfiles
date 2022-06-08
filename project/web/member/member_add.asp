<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include file = "htUIGen.inc" -->
<% 	xRootID=session("xRootID") %>
<%
'response.write  session("mySiteID")
cq = chr(34)
'response.write  session("mySiteID")
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

function htmlMessage(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	htmlMessage=""
  	exit function
  elseif instr(1,xs,"<P",1)>0 or instr(1,xs,"<BR",1)>0 or instr(1,xs,"<td",1)>0 then
  	htmlMessage=xs
  	exit function
  end if
  	xs = replace(xs,vbCRLF&vbCRLF,"<P>")
  	xs = replace(xs,vbCRLF,"<BR>")
  	htmlMessage = replace(xs,chr(10),"<BR>")
end function

function qqRS(xValue)
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
		qqRS = replace(xqqRS,chr(10),chr(34)&"&vbCRLF&"&chr(34))
	end if
end function 

	

	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
'		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\memberdata.xml"
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/wsxd2/member/xmlspec\memberdata.xml")
'		response.write LoadXML
'		response.end
		xv = htPageDom.load(LoadXML)
	  	if htPageDom.parseError.reason <> "" then 
	    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
	    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
	    		Response.End()
		end if
  	set session("hyXFormSpec") = htPageDom
  	set refModel = htPageDom.selectSingleNode("//dsTable")
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")
  	session("pickField") = true


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>加入會員</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--<script language="JavaScript" src="../js/mof.js" type="text/JavaScript"></script>-->
<link href="xslgip/mof94/css/member.css" rel="stylesheet" type="text/css" />

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<noscript>
  <p>本網頁使用script可是您的瀏覽器並不支援</p>
</noscript>
<!-- path start -->
<div class="path">
<a href="default.asp">首頁</a> &amp;gt; <a href="sp.asp?xdURL=member/member_add.asp">加入會員</a>
</div>
<!-- path end -->
<div class="contentword" >請填寫以下資料加入會員，＊部分為必填欄位。</div>
<table border="0" cellpadding="5" cellspacing="1"  summary="會員欄位" class="DataTb1" >
	<form name="reg" method="post" action="sp.asp?xdURL=member/member_add_act.asp">
	<!--form name="reg" method="post" action="xdsp.asp?xdURL=member/memberadd_act.asp&deptID=<%= deptID %>&epaperID=<%=epaperID%>&mp=1"-->
	<!--form name="reg" method="post" action="member/memberadd_act.asp?deptID=<%= deptID %>&epaperID=<%=epaperID%>"-->
<%	
	for each param in refModel.selectNodes("//fieldList/field[show='Y']") 
	    if nullText(param.selectSingleNode("inputType"))<>"hidden"then
		    response.write "<TR><TH align=""right"" width=""18%"" nowrap=""nowrap"">"
		    if nullText(param.selectSingleNode("canNull")) = "N" then _
				response.write "<span class=""Must"">*</span>"
		    response.write nullText(param.selectSingleNode("fieldLabel")) & "</TH>"
		    response.write "<TD >"
	    end if
		processParamField param	  
	    if nullText(param.selectSingleNode("inputType"))<>"hidden"then
			response.write "</TD></TR>"
		end if		
	next

%>

<INPUT TYPE=hidden name=CalendarTarget>
          <tr>
            <td >&amp;nbsp;</td>
            <td >
              <input name="submit.x" type="submit" id="submit" value="確定" width="44" height="17" border="0">              　
              <input name="cancel.x" type="submit" id="cancel" value="取消" width="44" height="17" border="0"></td>
          </tr>
	</form>

<script language=vbs>
cvbCRLF = vbCRLF
cTabchar = chr(9)

Dim CanTarget
Dim followCanTarget

sub popCalendar(dateName,followName)        
       ex=window.event.clientX
       ey=document.body.scrolltop+window.event.clientY+10
'       msgBox ex & "," & ey & " --xx " & document.body.scrolltop
       if ex>520 then ex=520

 	CanTarget=dateName
 	followCanTarget=followName
	xdate = document.all(CanTarget).value
	if not isDate(xdate) then	xdate = date()
	document.all.calendar1.setDate xdate
	
 	If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 	Else        
       document.all.calendar1.style.pixelleft=ex-80
       document.all.calendar1.style.pixeltop=ey
       document.all.calendar1.style.visibility=""
 	End If              
end sub  

sub calendar1_onscriptletevent(n,o) 
    	document.all("CalendarTarget").value=n         
    	document.all.calendar1.style.visibility="hidden"    
    if n <> "Cancle" then
        document.all(CanTarget).value=document.all.CalendarTarget.value
        document.all("pcShow"&CanTarget).value=document.all.CalendarTarget.value
        if followCanTarget<>"" then
        	document.all(followCanTarget).value=document.all.CalendarTarget.value
        	document.all("pcShow"&followCanTarget).value=document.all.CalendarTarget.value
        end if
	end if
end sub   

<%	for each param in refModel.selectNodes("//fieldList/field") 
		    AddProcessInit param
	next
%>
'	clientInitForm
 </script>	
</table>
</body>
</html>
