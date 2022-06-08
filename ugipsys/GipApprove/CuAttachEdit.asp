<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="޲z"
HTProgFunc="s"
HTUploadPath=session("Public")+"Attachment/"
HTProgCode="GC1AP2"
HTProgPrefix="CuAttach" %>
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
				if xup.IsFileExist( apath & xUpForm("hoFile_NFileName")) then _
					xup.DeleteFile apath & xUpForm("hoFile_NFileName")
'	response.write apath & xUpForm("hoFile_NFileName")
'	response.end
	SQL = "DELETE FROM CuDTAttach WHERE ixCuAttach=" & pkStr(request.queryString("ixCuAttach"),"")
	conn.Execute SQL
	showDoneBox("ƧR\I")

else
	EditInBothCase()
end if


sub EditInBothCase
	sqlCom = "SELECT htx.*, xrefNFileName.oldFileName AS fxr_NFileName FROM (CuDTAttach AS htx LEFT JOIN imageFile AS xrefNFileName ON xrefNFileName.newFileName = htx.NFileName) WHERE htx.ixCuAttach=" & pkStr(request.queryString("ixCuAttach"),"")
	Set RSreg = Conn.execute(sqlcom)
	pKey = ""
	pKey = pKey & "&ixCuAttach=" & RSreg("ixCuAttach")
	if pKey<>"" then  pKey = mid(pKey,2)

	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
end sub


function qqRS(fldName)
	if xUpForm("submitTask")="" then
		xValue = RSreg(fldname)
	else
		xValue = ""
		if xUpForm("htx_"&fldName) <> "" then
			xValue = xUpForm("htx_"&fldName)
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
	reg.htx_aTitle.value= "<%=qqRS("aTitle")%>"
	reg.htx_ixCuAttach.value= "<%=qqRS("ixCuAttach")%>"
	reg.htx_xiCuItem.value= "<%=qqRS("xiCuItem")%>"
	reg.htx_aDesc.value= "<%=qqRS("aDesc")%>"
	initAttFile "NFileName", "<%=qqRS("NFileName")%>", "<%=qqRS("fxr_NFileName")%>"
	reg.htx_bList.value= "<%=qqRS("bList")%>"
	reg.htx_listSeq.value= "<%=qqRS("listSeq")%>"
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
  
	IF reg.htx_aTitle.value = Empty Then 
		MsgBox replace(nMsg,"{0}","W"), 64, "Sorry!"
		reg.htx_aTitle.focus
		exit sub
	END IF
	IF blen(reg.htx_aTitle.value) > 80 Then
		MsgBox replace(replace(lMsg,"{0}","W"),"{1}","80"), 64, "Sorry!"
		reg.htx_aTitle.focus
		exit sub
	END IF
	IF (reg.htx_ixCuAttach.value <> "") AND (NOT isNumeric(reg.htx_ixCuAttach.value)) Then
		MsgBox replace(iMsg,"{0}","ID"), 64, "Sorry!"
		reg.htx_ixCuAttach.focus
		exit sub
	END IF
	IF (reg.htx_xiCuItem.value <> "") AND (NOT isNumeric(reg.htx_xiCuItem.value)) Then
		MsgBox replace(iMsg,"{0}","ID"), 64, "Sorry!"
		reg.htx_xiCuItem.focus
		exit sub
	END IF
	IF blen(reg.htx_aDesc.value) > 800 Then
		MsgBox replace(replace(lMsg,"{0}",""),"{1}","800"), 64, "Sorry!"
		reg.htx_aDesc.focus
		exit sub
	END IF
	IF reg.htfile_NFileName.value = Empty Then 
'		MsgBox replace(nMsg,"{0}","tɦW"), 64, "Sorry!"
'		reg.htfile_NFileName.focus
'		exit sub
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

<!--#include file="CuAttachFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>i<%=HTProgFunc%>j</td>
		<td width="50%" class="FormLink" align="right">
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
	sql = "UPDATE CuDTAttach SET "
		sql = sql & "aTitle=" & pkStr(xUpForm("htx_aTitle"),",")
		sql = sql & "aDesc=" & pkStr(xUpForm("htx_aDesc"),",")
		sql = sql & "bList=" & pkStr(xUpForm("htx_bList"),",")
		sql = sql & "listSeq=" & pkStr(xUpForm("htx_listSeq"),",")
	IF xUpForm("htFileActCK_NFileName") <> "" Then
	  actCK = xUpForm("htFileActCK_NFileName")
	  if actCK="editLogo" OR actCK="addLogo" then
		fname = ""
		For each xatt in xup.Attachments
		  if xatt.Name = "htFile_NFileName" then
			ofname = xatt.FileName
			fnExt = ""
			if instrRev(ofname, ".")>0 then	fnext=mid(ofname, instrRev(ofname, "."))
			tstr = now()
			nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
			sql = sql & "NFileName=" & pkStr(nfname,",")
			IF xUpForm("hoFile_NFileName") <> "" Then
				if xup.IsFileExist( apath & xUpForm("hoFile_NFileName")) then _
					xup.DeleteFile apath & xUpForm("hoFile_NFileName")
				xsql = "DELETE imageFile WHERE newFileName=" & pkStr(xUpForm("hoFile_NFileName"),"")
				conn.execute xsql
			END IF
			xatt.SaveFile apath & nfname, false
			xsql = "INSERT INTO imageFile(newFileName, oldFileName) VALUES(" & pkStr(nfname,",") & pkStr(ofname,")")
			conn.execute xsql
		  end if
		Next
	  elseif actCK="delLogo" then
		if xup.IsFileExist( apath & xUpForm("hoFile_NFileName")) then _
			xup.DeleteFile apath & xUpForm("hoFile_NFileName")
		xsql = "DELETE imageFile WHERE newFileName=" & pkStr(xUpForm("hoFile_NFileName"),"")
		conn.execute xsql
		sql = sql & "NFileName=null,"
	  end if
	END IF
	sql = left(sql,len(sql)-1) & " WHERE ixCuAttach=" & pkStr(request.queryString("ixCuAttach"),"")

	conn.Execute(SQL)  
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) 
	mpKey = ""
	mpKey = mpKey & "&iCuItem=" & xUpForm("htx_xiCuItem")
	if mpKey<>"" then  mpKey = mid(mpKey,2)
	doneURI= ""
	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp"
%>
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
	    document.location.href="<%=doneURI%>?<%=mpKey%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
