<% Response.Expires = 0
HTProgCap="�ҵ{�޲z"
HTProgFunc="�覸�s��"
HTProgCode="PA001"
HTProgPrefix="paSession" %>

<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
taskLable="�s��" & HTProgCap

if request("submitTask") = "UPDATE" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("��Ƨ�s���\�I")
	end if

elseif request("submitTask") = "DELETE" then
	SQL = "DELETE FROM paSession WHERE paSID=" & pkStr(request.queryString("paSID"),"")
	conn.Execute SQL
	showDoneBox("��ƧR�����\�I")

else
	EditInBothCase()
end if


sub EditInBothCase
	sqlCom = "SELECT * FROM paSession WHERE paSID=" & pkStr(request.queryString("paSID"),"")
	Set RSreg = Conn.execute(sqlcom)
	pKey = ""
	pKey = pKey & "&paSID=" & RSreg("paSID")
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
	reg.htx_bDate.value= "<%=qqRS("bDate")%>"
	reg.htx_actID.value= "<%=qqRS("actID")%>"
	reg.htx_paSID.value= "<%=qqRS("paSID")%>"
	reg.htx_dtNote.value= "<%=qqRS("dtNote")%>"
	reg.htx_pLimit.value= "<%=qqRS("pLimit")%>"
	reg.htx_paSnum.value= "<%=qqRS("paSnum")%>"
	reg.htx_aStatus.value= "<%=qqRS("aStatus")%>"
	reg.htx_refPage.value= "<%=qqRS("refPage")%>"
	reg.htx_place.value= "<%=qqRS("place")%>"
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
  
	IF reg.htx_bDate.value = Empty Then
		MsgBox replace(nMsg,"{0}","�ҵ{�_�l��"), 64, "Sorry!"
		reg.htx_bDate.focus
		exit sub
	END IF
	IF reg.htx_actID.value = Empty Then
		MsgBox replace(nMsg,"{0}","�ҵ{�N�X"), 64, "Sorry!"
		reg.htx_actID.focus
		exit sub
	END IF
	IF reg.htx_paSID.value = Empty Then
		MsgBox replace(nMsg,"{0}","�覸���X"), 64, "Sorry!"
		reg.htx_paSID.focus
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

<!--#include file="paSessionFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;"> 
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
    <title>�s�ת���</title>
    <link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
    </head>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>�i<%=HTProgFunc%>�j</td>
		<td width="50%" class="FormLink" align="right">
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <a href="Javascript:window.history.back();">�^�e��</a>
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
	sql = "UPDATE paSession SET "
		sql = sql & "bDate=" & pkStr(request("htx_bDate"),",")
		sql = sql & "actID=" & pkStr(request("htx_actID"),",")
		sql = sql & "dtNote=" & pkStr(request("htx_dtNote"),",")
		sql = sql & "pLimit=" & pkStr(request("htx_pLimit"),",")
		sql = sql & "paSnum=" & pkStr(request("htx_paSnum"),",")
		sql = sql & "aStatus=" & pkStr(request("htx_aStatus"),",")
		sql = sql & "refPage=" & pkStr(request("htx_refPage"),",")
		sql = sql & "place=" & pkStr(request("htx_place"),",")
	sql = left(sql,len(sql)-1) & " WHERE paSID=" & pkStr(request.queryString("paSID"),"")

	conn.Execute(SQL)  
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) 
	mpKey = ""
	mpKey = mpKey & "&actID=" & request("htx_actID")
	if mpKey<>"" then  mpKey = mid(mpKey,2)
%>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=big5">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>�s�ת���</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
	    document.location.href="<%=HTprogPrefix%>List.asp?<%=mpKey%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>