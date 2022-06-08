<% Response.Expires = 0
HTProgCap="�ҵ{�޲z"
HTProgFunc="�s�W"
HTProgCode="PA001"
HTProgPrefix="paAct" %>
<!--#Include Virtual="inc/server.inc" -->
<!--#INCLUDE Virtual="inc/dbutil.inc" -->
<%
if request("submitTask") = "ADD" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		showErrBox()
	else
		doUpdateDB()
		showDoneBox()
	end if

else

	showHTMLHead()
	formFunction = "add"
	showForm()
	initForm()
	showHTMLTail()
end if
%>

<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm	'---- �s�W�ɪ����w�]�ȩ�b�o��
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

    sub initCheckbox(xname,ckValue)
    	value = ckValue & ","
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then
    			xri.checked = true
    		end if
    	next
    end sub
</script>
<% end sub '---- initForm() ----%>

<% sub showForm() 	'===================== Client Side Validation Put HERE =========== %>
<script language="vbscript">

Sub formAddSubmit()

'----- �Τ��Submit�e���ˬd�X��b�o�� �A�p�U�� not valid �� exit sub ���}------------
'msg1 = "�аȥ���g�u�Ȥ�s���v�A���o���ťաI"
'msg2 = "�аȥ���g�u�Ȥ�W�١v�A���o���ťաI"
'
'  If reg.pfx_ClientID.value = Empty Then
'     MsgBox msg1, 64, "Sorry!"
'     settab 0
'     reg.pfx_ClientID.focus
'     Exit Sub
'  End if
	nMsg = "�аȥ���g�u{0}�v�A���o���ťաI"
	lMsg = "�u{0}�v�����׳̦h��{1}�I"
	dMsg = "�u{0}�v������� yyyy/mm/dd �I"
	iMsg = "�u{0}�v��������ƭȡI"
	pMsg = "�u{0}�v���������������� " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" �䤤�@�ءI"
	IF reg.htx_actCat.value = Empty Then
		MsgBox replace(nMsg,"{0}","�ҵ{���O"), 64, "Sorry!"
		reg.htx_actCat.focus
		exit sub
	END IF
	IF reg.htx_actName.value = Empty Then
		MsgBox replace(nMsg,"{0}","���ئW��"), 64, "Sorry!"
		reg.htx_actName.focus
		exit sub
	END IF
	IF reg.htx_actDesc.value = Empty Then
		MsgBox replace(nMsg,"{0}","���e"), 64, "Sorry!"
		reg.htx_actDesc.focus
		exit sub
	END IF
	IF blen(reg.htx_actDesc.value) > 1000 Then
		MsgBox replace(replace(lMsg,"{0}","���e"),"{1}","1000"), 64, "Sorry!"
		reg.htx_actDesc.focus
		exit sub
	END IF
	IF reg.htx_actTarget.value = Empty Then
		MsgBox replace(nMsg,"{0}","��H"), 64, "Sorry!"
		reg.htx_actTarget.focus
		exit sub
	END IF
	IF reg.htx_actTime.value = Empty Then
		MsgBox replace(nMsg,"{0}","�ɶ�"), 64, "Sorry!"
		reg.htx_actTime.focus
		exit sub
	END IF

  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="paActForm.inc"-->

<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>�s�W���</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>�i<%=HTProgFunc%>�j</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <%if (HTProgRight and 2)=2 then%>
	       <a href="<%=HTprogPrefix%>Query.asp">�d��</a>
	       <%end if%>
		</td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>
<% end sub '--- showHTMLHead() ------%>


<% sub ShowHTMLTail() %>
    </td>
  </tr>
</table>
</body>
</html>
<% end sub '--- showHTMLTail() ------%>


<% sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
  			window.history.back
	</script>
<%
end sub '---- showErrBox() ----

sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- ��ݸ�������ˬd�{���X��b�o��	�A�p�U�ҡA�����ɳ] errMsg="xxx" �� exit sub ------
'	SQL = "Select * From Client Where ClientID = '"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "�u�Ȥ�s���v����!!�Э��s�إ߫Ȥ�s��!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

sub doUpdateDB()
	sql = "INSERT INTO ppAct("
	sqlValue = ") VALUES("
	IF request("htx_actCat") <> "" Then
		sql = sql & "actCat" & ","
		sqlValue = sqlValue & pkStr(request("htx_actCat"),",")
	END IF
	IF request("htx_actName") <> "" Then
		sql = sql & "actName" & ","
		sqlValue = sqlValue & pkStr(request("htx_actName"),",")
	END IF
	IF request("htx_actDesc") <> "" Then
		sql = sql & "actDesc" & ","
		sqlValue = sqlValue & pkStr(request("htx_actDesc"),",")
	END IF
	IF request("htx_actTarget") <> "" Then
		sql = sql & "actTarget" & ","
		sqlValue = sqlValue & pkStr(request("htx_actTarget"),",")
	END IF
	IF request("htx_detailDesc") <> "" Then
		sql = sql & "detailDesc" & ","
		sqlValue = sqlValue & pkStr(request("htx_detailDesc"),",")
	END IF
	IF request("htx_place") <> "" Then
		sql = sql & "place" & ","
		sqlValue = sqlValue & pkStr(request("htx_place"),",")
	END IF
	IF request("htx_actTime") <> "" Then
		sql = sql & "actTime" & ","
		sqlValue = sqlValue & pkStr(request("htx_actTime"),",")
	END IF
	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"

	conn.Execute(SQL)
end sub '---- doUpdateDB() ----

%>

<% sub showDoneBox() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>�s�W�{��</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	alert("�s�W�����I")
	document.location.href = "<%=HTprogPrefix%>List.asp"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
