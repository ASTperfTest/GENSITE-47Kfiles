<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="ql޲z"
HTProgFunc="s׵o"
HTUploadPath="/public/"
HTProgCode="GW1M51"
HTProgPrefix="ePub" %>
<!--#include virtual = "/inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
<title>sת</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
taskLable="s" & HTProgCap

 apath=server.mappath(HTUploadPath) & "\"
  set xup = Server.CreateObject("UpDownExpress.FileUpload")
  xup.Open 

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

if xUpForm("submitTask") = "UPDATE" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("Ƨs\I")
	end if

elseif xUpForm("submitTask") = "DELETE" then
	SQL = "DELETE FROM EpPub WHERE ePubID=" & pkStr(request.queryString("ePubID"),"")
	conn.Execute SQL
	showDoneBox("ƧR\I")

else
	EditInBothCase()
end if


sub EditInBothCase
	sqlCom = "SELECT htx.* FROM EpPub AS htx WHERE htx.ePubID=" & pkStr(request.queryString("ePubID"),"")
	Set RSreg = Conn.execute(sqlcom)
	pKey = ""
	pKey = pKey & "&ePubID=" & RSreg("ePubID")
	if pKey<>"" then  pKey = mid(pKey,2)

	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
end sub


function qqRS(fldName)
	if request("submitTask")="" then
		xValue = RSreg(fldname)
	else
		xValue = ""
		if request("htx_"&fldName) <> "" then
			xValue = request("htx_"&fldName)
		end if
	end if
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end if
end function %>

<%Sub initForm() %>
    <script language=vbs>
      sub resetForm 
	       reg.reset()
	       clientInitForm
      end sub

     sub clientInitForm
	reg.htx_pubDate.value= "<%=qqRS("pubDate")%>"
	reg.pcShowhtx_pubDate.value= "<%=qqRS("pubDate")%>"
	reg.htx_ePubID.value= "<%=qqRS("ePubID")%>"
	reg.htx_title.value= "<%=qqRS("title")%>"
	reg.htx_dbDate.value= "<%=qqRS("dbDate")%>"
	reg.pcShowhtx_dbDate.value= "<%=qqRS("dbDate")%>"
	reg.htx_deDate.value= "<%=qqRS("deDate")%>"
	reg.pcShowhtx_deDate.value= "<%=qqRS("deDate")%>"
	reg.htx_maxNo.value= "<%=qqRS("maxNo")%>"
    end sub

    sub window_onLoad
         clientInitForm
    end sub
    
    sub initRadio(xname,value)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    end sub

    sub initOtherRadio(xname,value, otherName)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    	if value="" then	exit sub
		reg.all(xname).item(reg.all(xname).length-1).checked = true
		reg.all(otherName).value = value
		reg.all(xname).item(reg.all(xname).length-1).value = value
    end sub

    sub initCheckbox(xname,ckValue)
    	value = ckValue & ","
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    		end if
    	next
    end sub
    
    sub initOtherCheckbox(xname,ckValue,otherName)
    	valueArray = split(ckValue,", ")
    	valueCount = ubound(valueArray) + 1
    	value = ckValue & ","
    	ckCount = 0
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    			ckCount = ckCount + 1
    		end if
    	next
		if ckCount <> valueCount then
			reg.all(xname).item(reg.all(xname).length-1).checked = true
			reg.all(otherName).value = valueArray(ubound(valueArray))
			reg.all(xname).item(reg.all(xname).length-1).value = valueArray(ubound(valueArray))
		end if
    end sub

    sub initImgFile(xname, value)
		reg.all("htImgActCK_"&xname).value=""
		reg.all("htImg_"&xname).style.display="none"
		reg.all("hoImg_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).src= "<%=HTUploadPath%>" & value
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub

Sub addLogo(xname)	'sWlogo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="addLogo"
End sub

Sub chgLogo(xname)	'logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="editLogo"
End sub

Sub orgLogo(xname)	'
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value=""
End sub

Sub delLogo(xname)	'Rlogo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htImg_"&xname).value=""
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value="delLogo"
End sub

    sub initAttFile(xname, value, orgValue)
		reg.all("htFileActCK_"&xname).value=""
		reg.all("htFile_"&xname).style.display="none"
		reg.all("hoFile_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).innerText= orgValue
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub

Sub addXFile(xname)	'sWlogo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="addLogo"
End sub

Sub chgXFile(xname)	'logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="editLogo"
End sub

Sub orgXFile(xname)	'
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value=""
End sub

Sub delXFile(xname)	'Rlogo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htFile_"&xname).value=""
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value="delLogo"
End sub

 </script>	
<%End sub '---- initForm() ----%>


<%Sub showForm()  	'===================== Client Side Validation Put HERE =========== %>
    <script language="vbscript">
      sub formModSubmit()
    
'----- ΤSubmiteˬdXbo ApU not valid  exit sub }------------
'msg1 = "аȥguȤsvAoťաI"
'msg2 = "аȥguȤW١vAoťաI"
'
'  If reg.tfx_ClientName.value = Empty Then
'     MsgBox msg2, 64, "Sorry!"
'     settab 0
'     reg.tfx_ClientName.focus
'     Exit Sub
'  End if
	nMsg = "аȥgu{0}vAoťաI"
	lMsg = "u{0}v׳̦h{1}I"
	dMsg = "u{0}v yyyy/mm/dd I"
	iMsg = "u{0}vƭȡI"
	pMsg = "u{0}v " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 䤤@ءI"
  
	IF reg.htx_pubDate.value = Empty Then 
		MsgBox replace(nMsg,"{0}","o"), 64, "Sorry!"
		reg.htx_pubDate.focus
		exit sub
	END IF
	IF (reg.htx_pubDate.value <> "") AND (NOT isDate(reg.htx_pubDate.value)) Then
		MsgBox replace(dMsg,"{0}","o"), 64, "Sorry!"
		reg.htx_pubDate.focus
		exit sub
	END IF
	IF (reg.htx_ePubID.value <> "") AND (NOT isNumeric(reg.htx_ePubID.value)) Then
		MsgBox replace(iMsg,"{0}","oID"), 64, "Sorry!"
		reg.htx_ePubID.focus
		exit sub
	END IF
	IF reg.htx_title.value = Empty Then 
		MsgBox replace(nMsg,"{0}","D"), 64, "Sorry!"
		reg.htx_title.focus
		exit sub
	END IF
	IF blen(reg.htx_title.value) > 80 Then
		MsgBox replace(replace(lMsg,"{0}","D"),"{1}","80"), 64, "Sorry!"
		reg.htx_title.focus
		exit sub
	END IF
	IF reg.htx_dbDate.value = Empty Then 
		MsgBox replace(nMsg,"{0}","ƽd_"), 64, "Sorry!"
		reg.htx_dbDate.focus
		exit sub
	END IF
	IF (reg.htx_dbDate.value <> "") AND (NOT isDate(reg.htx_dbDate.value)) Then
		MsgBox replace(dMsg,"{0}","ƽd_"), 64, "Sorry!"
		reg.htx_dbDate.focus
		exit sub
	END IF
	IF reg.htx_deDate.value = Empty Then 
		MsgBox replace(nMsg,"{0}","ƽd򨴤"), 64, "Sorry!"
		reg.htx_deDate.focus
		exit sub
	END IF
	IF (reg.htx_deDate.value <> "") AND (NOT isDate(reg.htx_deDate.value)) Then
		MsgBox replace(dMsg,"{0}","ƽd򨴤"), 64, "Sorry!"
		reg.htx_deDate.focus
		exit sub
	END IF
	IF reg.htx_maxNo.value = Empty Then 
		MsgBox replace(nMsg,"{0}","ƫh"), 64, "Sorry!"
		reg.htx_maxNo.focus
		exit sub
	END IF
	IF (reg.htx_maxNo.value <> "") AND (NOT isNumeric(reg.htx_maxNo.value)) Then
		MsgBox replace(iMsg,"{0}","ƫh"), 64, "Sorry!"
		reg.htx_maxNo.focus
		exit sub
	END IF
  
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub

   sub formDelSubmit()
         chky=msgbox("`NI"& vbcrlf & vbcrlf &"@ATwRƶܡH"& vbcrlf , 48+1, "Ъ`NII")
        if chky=vbok then
	       reg.submitTask.value = "DELETE"
	      reg.Submit
       end If
  end sub

</script>

<!--#include file="ePubFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>i<%=HTProgFunc%>--<%=session("ePaperName")%>j</td>
		<td width="50%" class="FormLink" align="right">
		<%if (HTProgRight and 8)=8 then%>
			<A href="ePubSend.asp?ePubID=<%=request.queryString("ePubID")%>" title="oe">oe</A>
			<a href="JavaScript:ePubGen();" title="ͤe">ͤe</a>
		<%end if%>
			<A href="Javascript:window.history.back();" title="^e">^e</A> 
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>        
<%End sub '--- showHTMLHead() ------%>


<%Sub ShowHTMLTail() %>     
    </td>
  </tr>  
</table> 
</body>
</html>
<%End sub '--- showHTMLTail() ------%>


<%Sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
'  			window.history.back
	</script>
<%
    End sub '---- showErrBox() ----

Sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- ݸˬd{Xbo	ApUҡAɳ] errMsg="xxx"  exit sub ------
'	if Request("tfx_TaxID") <> "" Then
'	  SQL = "Select * From Client Where TaxID = N'"& Request("tfx_TaxID") &"'"
'	  set RSreg = conn.Execute(SQL)
'	  if not RSreg.EOF then
'		if trim(RSreg("ClientId")) <> request("pfx_ClientID") Then
'			errMsg = "uΤ@sv!!ЭsJȤΤ@s!"
'			exit sub
'		end if
'	  end if
'	end if

End sub '---- checkDBValid() ----

Sub doUpdateDB()
	sql = "UPDATE EpPub SET "
		sql = sql & "pubDate=" & pkStr(xUpForm("htx_pubDate"),",")
		sql = sql & "title=" & pkStr(xUpForm("htx_title"),",")
		sql = sql & "dbDate=" & pkStr(xUpForm("htx_dbDate"),",")
		sql = sql & "deDate=" & pkStr(xUpForm("htx_deDate"),",")
		sql = sql & "maxNo=" & pkStr(xUpForm("htx_maxNo"),",")
	sql = left(sql,len(sql)-1) & " WHERE ePubID=" & pkStr(request.queryString("ePubID"),"")

	conn.Execute(SQL)  
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>sת</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
	    document.location.href="<%=HTprogPrefix%>List.asp?<%=Session("QueryPage_No")%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
<script language=vbs>
sub ePubGen()
     	window.open "ePubGen.asp?ePubID=<%=request.queryString("ePubID")%>"
end sub

Dim CanTarget
Dim followCanTarget

sub popCalendar(dateName,followName)        
 	CanTarget=dateName
 	followCanTarget=followName
	xdate = document.all(CanTarget).value
	if not isDate(xdate) then	xdate = date()
	document.all.calendar1.setDate xdate
	
 	If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 	Else        
       ex=window.event.clientX
       ey=document.body.scrolltop+window.event.clientY+10
       if ex>520 then ex=520
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
</script>
