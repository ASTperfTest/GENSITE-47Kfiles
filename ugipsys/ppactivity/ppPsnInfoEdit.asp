<% Response.Expires = 0
HTProgCap="���ʾǭ��޲z"
HTProgFunc="�s��"
HTUploadPath="/public/"
HTProgCode="PA010"
HTProgPrefix="ppPsnInfo" %>
<!--#Include file = "../inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5; no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
<title>�s�ת��</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
taskLable="�s��" & HTProgCap

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
		showDoneBox("��Ƨ�s���\�I")
	end if

elseif xUpForm("submitTask") = "DELETE" then
	SQL = "DELETE FROM paPsnInfo WHERE psnID=" & pkStr(request.queryString("psnID"),"")
	conn.Execute SQL
	showDoneBox("��ƧR�����\�I")

else
	EditInBothCase()
end if


sub EditInBothCase
	sqlCom = "SELECT htx.* FROM paPsnInfo AS htx WHERE htx.psnID=" & pkStr(request.queryString("psnID"),"")
	Set RSreg = Conn.execute(sqlcom)
	pKey = ""
	pKey = pKey & "&psnID=" & RSreg("psnID")
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
	reg.htx_myPassword.value= "<%=qqRS("myPassword")%>"
	reg.htx_myPassword2check.value= "<%=qqRS("myPassword")%>"
	reg.htx_psnID.value= "<%=qqRS("psnID")%>"
	reg.htx_pName.value= "<%=qqRS("pName")%>"
	reg.htx_birthDay.value= "<%=qqRS("birthDay")%>"
	reg.pcShowhtx_birthDay.value= d7Date("<%=qqRS("birthDay")%>")
	reg.htx_eMail.value= "<%=qqRS("eMail")%>"
	reg.htx_tel.value= "<%=qqRS("tel")%>"
	reg.htx_emergContact.value= "<%=qqRS("emergContact")%>"
	reg.htx_deptName.value= "<%=qqRS("deptName")%>"
	reg.htx_jobName.value= "<%=qqRS("jobName")%>"
	reg.htx_topEdu.value= "<%=qqRS("topEdu")%>"
	reg.htx_meatKind.value= "<%=qqRS("meatKind")%>"
	reg.htx_corpName.value= "<%=qqRS("corpName")%>"
	reg.htx_corpID.value= "<%=qqRS("corpID")%>"
	reg.htx_corpAddr.value= "<%=qqRS("corpAddr")%>"
	reg.htx_corpTel.value= "<%=qqRS("corpTel")%>"
	reg.htx_corpFax.value= "<%=qqRS("corpFax")%>"
	reg.htx_datafrom.value= "<%=qqRS("datafrom")%>"
	reg.htx_mobile.value= "<%=qqRS("mobile")%>"
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

Sub addLogo(xname)	'�s�Wlogo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="addLogo"
End sub

Sub chgLogo(xname)	'��logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="editLogo"
End sub

Sub orgLogo(xname)	'���
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value=""
End sub

Sub delLogo(xname)	'�R��logo
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

Sub addXFile(xname)	'�s�Wlogo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="addLogo"
End sub

Sub chgXFile(xname)	'��logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="editLogo"
End sub

Sub orgXFile(xname)	'���
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value=""
End sub

Sub delXFile(xname)	'�R��logo
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
    
'----- �Τ��Submit�e���ˬd�X��b�o�� �A�p�U�� not valid �� exit sub ���}------------
'msg1 = "�аȥ���g�u�Ȥ�s���v�A���o���ťաI"
'msg2 = "�аȥ���g�u�Ȥ�W�١v�A���o���ťաI"
'
'  If reg.tfx_ClientName.value = Empty Then
'     MsgBox msg2, 64, "Sorry!"
'     settab 0
'     reg.tfx_ClientName.focus
'     Exit Sub
'  End if
	nMsg = "�аȥ���g�u{0}�v�A���o���ťաI"
	lMsg = "�u{0}�v�����׳̦h��{1}�I"
	dMsg = "�u{0}�v������� yyyy/mm/dd �I"
	iMsg = "�u{0}�v��������ƭȡI"
	pMsg = "�u{0}�v���������������� " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" �䤤�@�ءI"
  
	IF reg.htx_myPassword.value = Empty Then 
		MsgBox replace(nMsg,"{0}","�n�J�K�X"), 64, "Sorry!"
		reg.htx_myPassword.focus
		exit sub
	END IF
	IF reg.htx_myPassword.value <>  reg.htx_myPassword2check.value  Then 
		MsgBox "�K�X��P�T�{�檺��J��ƻݬۦP", 64, "Sorry!"
		reg.htx_myPassword.focus
		exit sub
	END IF
	IF reg.htx_psnID.value = Empty Then 
		MsgBox replace(nMsg,"{0}","�����Ҹ�"), 64, "Sorry!"
		reg.htx_psnID.focus
		exit sub
	END IF
	IF blen(reg.htx_psnID.value) > 10 Then
		MsgBox replace(replace(lMsg,"{0}","�����Ҹ�"),"{1}","10"), 64, "Sorry!"
		reg.htx_psnID.focus
		exit sub
	END IF
	IF reg.htx_pName.value = Empty Then 
		MsgBox replace(nMsg,"{0}","�m�W"), 64, "Sorry!"
		reg.htx_pName.focus
		exit sub
	END IF
	IF blen(reg.htx_pName.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","�m�W"),"{1}","20"), 64, "Sorry!"
		reg.htx_pName.focus
		exit sub
	END IF
	IF reg.htx_birthDay.value = Empty Then 
		MsgBox replace(nMsg,"{0}","�X�ͤ�"), 64, "Sorry!"
		reg.htx_birthDay.focus
		exit sub
	END IF
	IF (reg.htx_birthDay.value <> "") AND (NOT isDate(reg.htx_birthDay.value)) Then
		MsgBox replace(dMsg,"{0}","�X�ͤ�"), 64, "Sorry!"
		reg.htx_birthDay.focus
		exit sub
	END IF
	IF blen(reg.htx_eMail.value) > 50 Then
		MsgBox replace(replace(lMsg,"{0}","eMail"),"{1}","50"), 64, "Sorry!"
		reg.htx_eMail.focus
		exit sub
	END IF
	IF reg.htx_tel.value = Empty Then 
		MsgBox replace(nMsg,"{0}","�s���q��"), 64, "Sorry!"
		reg.htx_tel.focus
		exit sub
	END IF
	IF blen(reg.htx_tel.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","�s���q��"),"{1}","20"), 64, "Sorry!"
		reg.htx_tel.focus
		exit sub
	END IF
	IF reg.htx_emergContact.value = Empty Then 
		MsgBox replace(nMsg,"{0}","�p���a�}"), 64, "Sorry!"
		reg.htx_emergContact.focus
		exit sub
	END IF
	IF blen(reg.htx_emergContact.value) > 60 Then
		MsgBox replace(replace(lMsg,"{0}","�p���a�}"),"{1}","60"), 64, "Sorry!"
		reg.htx_emergContact.focus
		exit sub
	END IF
	IF blen(reg.htx_deptName.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","�u�@����"),"{1}","20"), 64, "Sorry!"
		reg.htx_deptName.focus
		exit sub
	END IF
	IF blen(reg.htx_jobName.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","¾��"),"{1}","20"), 64, "Sorry!"
		reg.htx_jobName.focus
		exit sub
	END IF
	IF blen(reg.htx_corpName.value) > 60 Then
		MsgBox replace(replace(lMsg,"{0}","���q����"),"{1}","60"), 64, "Sorry!"
		reg.htx_corpName.focus
		exit sub
	END IF
	IF blen(reg.htx_corpID.value) > 8 Then
		MsgBox replace(replace(lMsg,"{0}","�Τ@�s��"),"{1}","8"), 64, "Sorry!"
		reg.htx_corpID.focus
		exit sub
	END IF
	
	IF blen(reg.htx_corpAddr.value) > 60 Then
		MsgBox replace(replace(lMsg,"{0}","�p���a�}"),"{1}","60"), 64, "Sorry!"
		reg.htx_corpAddr.focus
		exit sub
	END IF
	IF blen(reg.htx_corpContact.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","�p���H"),"{1}","20"), 64, "Sorry!"
		reg.htx_corpContact.focus
		exit sub
	END IF
	IF blen(reg.htx_corpTel.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","�q��"),"{1}","20"), 64, "Sorry!"
		reg.htx_corpTel.focus
		exit sub
	END IF
	IF blen(reg.htx_corpFax.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","�ǯu"),"{1}","20"), 64, "Sorry!"
		reg.htx_corpFax.focus
		exit sub
	END IF
  
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub

   sub formDelSubmit()
         chky=msgbox("�`�N�I"& vbcrlf & vbcrlf &"�@�A�T�w�R����ƶܡH"& vbcrlf , 48+1, "�Ъ`�N�I�I")
        if chky=vbok then
	       reg.submitTask.value = "DELETE"
	      reg.Submit
       end If
  end sub

</script>

<!--#include file="ppPsnInfoFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>�i<%=HTProgFunc%>�j</td>
		<td width="50%" class="FormLink" align="right">
		<%if (HTProgRight and 1)=1 then%>
			<A href="ppPsnInfoQuery.asp" title="���]�d�߱���">�d��</A>
		<%end if%>
		<%if (HTProgRight and 4)=4 then%>
			<A href="ppPsnInfoAdd.asp" title="�s�W">�s�W</A>
		<%end if%>
			<A href="Javascript:window.history.back();" title="�^�e��">�^�e��</A> 
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

'---- ��ݸ�������ˬd�{���X��b�o��	�A�p�U�ҡA�����ɳ] errMsg="xxx" �� exit sub ------
'	if Request("tfx_TaxID") <> "" Then
'	  SQL = "Select * From Client Where TaxID = '"& Request("tfx_TaxID") &"'"
'	  set RSreg = conn.Execute(SQL)
'	  if not RSreg.EOF then
'		if trim(RSreg("ClientId")) <> request("pfx_ClientID") Then
'			errMsg = "�u�Τ@�s���v����!!�Э��s��J�Ȥ�Τ@�s��!"
'			exit sub
'		end if
'	  end if
'	end if

End sub '---- checkDBValid() ----

Sub doUpdateDB()
	sql = "UPDATE paPsnInfo SET "
		sql = sql & "myPassword=" & pkStr(xUpForm("htx_myPassword"),",")
		sql = sql & "psnID=" & pkStr(xUpForm("htx_psnID"),",")
		sql = sql & "pName=" & pkStr(xUpForm("htx_pName"),",")
		sql = sql & "birthDay=" & pkStr(xUpForm("htx_birthDay"),",")
		sql = sql & "eMail=" & pkStr(xUpForm("htx_eMail"),",")
		sql = sql & "tel=" & pkStr(xUpForm("htx_tel"),",")
		sql = sql & "emergContact=" & pkStr(xUpForm("htx_emergContact"),",")
		sql = sql & "deptName=" & pkStr(xUpForm("htx_deptName"),",")
		sql = sql & "jobName=" & pkStr(xUpForm("htx_jobName"),",")
		sql = sql & "topEdu=" & pkStr(xUpForm("htx_topEdu"),",")
		sql = sql & "meatKind=" & pkStr(xUpForm("htx_meatKind"),",")
		sql = sql & "corpName=" & pkStr(xUpForm("htx_corpName"),",")
		sql = sql & "corpID=" & pkStr(xUpForm("htx_corpID"),",")
		sql = sql & "corpAddr=" & pkStr(xUpForm("htx_corpAddr"),",")
		sql = sql & "corpContact=" & pkStr(xUpForm("htx_corpContact"),",")
		sql = sql & "corpTel=" & pkStr(xUpForm("htx_corpTel"),",")
		sql = sql & "corpFax=" & pkStr(xUpForm("htx_corpFax"),",")
		sql = sql & "datafrom=" & pkStr(xUpForm("htx_datafrom"),",")
		sql = sql & "mobile=" & pkStr(xUpForm("htx_mobile"),",")
	sql = left(sql,len(sql)-1) & " WHERE psnID=" & pkStr(request.queryString("psnID"),"")

	conn.Execute(SQL)  
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=big5">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>�s�ת��</title>
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
