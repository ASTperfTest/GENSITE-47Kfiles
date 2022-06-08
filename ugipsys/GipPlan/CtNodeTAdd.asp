<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="޲z"
HTProgFunc="sW`I"
HTUploadPath="/public/"
HTProgCode="GE1T21"
HTProgPrefix="CtNodeT" %>
<!--#include virtual = "/inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>sW</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
	dim xNewIdentity
 apath=server.mappath(HTUploadPath) & "\"
' response.write apath & "<HR>"
  set xup = Server.CreateObject("UpDownExpress.FileUpload")
  xup.Open 

if xUpForm("submitTask") = "ADD" then

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
<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm	'---- sWɪw]ȩbo
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
    
'----- ΤSubmiteˬdXbo ApU not valid  exit sub }------------
'msg1 = "аȥguȤsvAoťաI"
'msg2 = "аȥguȤW١vAoťաI"
'
'  If reg.pfx_ClientID.value = Empty Then
'     MsgBox msg1, 64, "Sorry!"
'     settab 0
'     reg.pfx_ClientID.focus
'     Exit Sub
'  End if
	nMsg = "аȥgu{0}vAoťաI"
	lMsg = "u{0}v׳̦h{1}I"
	dMsg = "u{0}v yyyy/mm/dd I"
	iMsg = "u{0}vƭȡI"
	pMsg = "u{0}v " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 䤤@ءI"
	IF reg.htx_CtNodeKind.value = Empty Then 
		MsgBox replace(nMsg,"{0}","`IO"), 64, "Sorry!"
		reg.htx_CtNodeKind.focus
		exit sub
	END IF
	IF (reg.htx_CtRootID.value <> "") AND (NOT isNumeric(reg.htx_CtRootID.value)) Then
		MsgBox replace(iMsg,"{0}","ID"), 64, "Sorry!"
		reg.htx_CtRootID.focus
		exit sub
	END IF
	IF (reg.htx_nodeParent.value <> "") AND (NOT isNumeric(reg.htx_nodeParent.value)) Then
		MsgBox replace(iMsg,"{0}","N`I"), 64, "Sorry!"
		reg.htx_nodeParent.focus
		exit sub
	END IF
	IF (reg.htx_CtUnitID.value <> "") AND (NOT isNumeric(reg.htx_CtUnitID.value)) Then
		MsgBox replace(iMsg,"{0}","DD椸ID"), 64, "Sorry!"
		reg.htx_CtUnitID.focus
		exit sub
	END IF
	IF (reg.htx_EditDate.value <> "") AND (NOT isDate(reg.htx_EditDate.value)) Then
		MsgBox replace(dMsg,"{0}","s"), 64, "Sorry!"
		reg.htx_EditDate.focus
		exit sub
	END IF
	IF (reg.htx_CtNodeID.value <> "") AND (NOT isNumeric(reg.htx_CtNodeID.value)) Then
		MsgBox replace(iMsg,"{0}","`IID"), 64, "Sorry!"
		reg.htx_CtNodeID.focus
		exit sub
	END IF
	IF reg.htx_CatName.value = Empty Then 
		MsgBox replace(nMsg,"{0}","D"), 64, "Sorry!"
		reg.htx_CatName.focus
		exit sub
	END IF
	IF blen(reg.htx_CatName.value) > 30 Then
		MsgBox replace(replace(lMsg,"{0}","D"),"{1}","30"), 64, "Sorry!"
		reg.htx_CatName.focus
		exit sub
	END IF
	IF reg.htImg_CtNameLogo.value <> "" Then
		xIMGname = reg.htImg_CtNameLogo.value
		xFileType = ""
		if instr(xIMGname, ".")>0 then	xFileType=lcase(mid(xIMGname, instr(xIMGname, ".")+1))
		IF xFileType<>"gif" AND xFileType<>"jpg" AND xFileType<>"jpeg" then
			MsgBox replace(pMsg,"{0}","Dϥ"), 64, "Sorry!"
			reg.htImg_CtNameLogo.focus
			exit sub
		END IF
	END IF
	IF reg.htx_inUse.value = Empty Then 
		MsgBox replace(nMsg,"{0}","O_}"), 64, "Sorry!"
		reg.htx_inUse.focus
		exit sub
	END IF
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="CtNodeTForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
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
	    <td align=center colspan=2 width=90% height=230 valign=top>    
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

'---- ݸˬd{Xbo	ApUҡAɳ] errMsg="xxx"  exit sub ------
'	SQL = "Select * From Client Where ClientID = N'"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "uȤsv!!Эsإ߫Ȥs!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

sub doUpdateDB()
	sql = "INSERT INTO CatTreeNode("
	sqlValue = ") VALUES("
	IF xUpForm("htx_CtNodeKind") <> "" Then
		sql = sql & "CtNodeKind" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_CtNodeKind"),",")
	END IF
	IF xUpForm("htx_CtRootID") <> "" Then
		sql = sql & "CtRootID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_CtRootID"),",")
	END IF
	IF xUpForm("htx_DataLevel") <> "" Then
		sql = sql & "DataLevel" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_DataLevel"),",")
	END IF
	IF xUpForm("htx_nodeParent") <> "" Then
		sql = sql & "nodeParent" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_nodeParent"),",")
	END IF
	IF xUpForm("htx_CtUnitID") <> "" Then
		sql = sql & "CtUnitID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_CtUnitID"),",")
	END IF
	IF xUpForm("htx_EditDate") <> "" Then
		sql = sql & "EditDate" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_EditDate"),",")
	END IF
	IF xUpForm("htx_EditUserID") <> "" Then
		sql = sql & "EditUserID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_EditUserID"),",")
	END IF
	IF xUpForm("htx_CtNodeID") <> "" Then
		sql = sql & "CtNodeID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_CtNodeID"),",")
	END IF
	IF xUpForm("htx_CatName") <> "" Then
		sql = sql & "CatName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_CatName"),",")
	END IF
	IF xUpForm("htx_CatShowOrder") <> "" Then
		sql = sql & "CatShowOrder" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_CatShowOrder"),",")
	END IF
	IF xUpForm("htx_inUse") <> "" Then
		sql = sql & "inUse" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_inUse"),",")
	END IF

  For each xatt in xup.Attachments
    if left(xatt.Name,6) = "htImg_" then
	  ofname = xatt.FileName
	  fnExt = ""
	  if instr(ofname, ".")>0 then	fnext=mid(ofname, instr(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(xatt.Name,7) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xatt.SaveFile apath & nfname, false
  	elseif left(xatt.Name,7) = "htFile_" then
	  ofname = xatt.FileName
	  fnExt = ""
	  if instr(ofname, ".")>0 then	fnext=mid(ofname, instr(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(xatt.Name,8) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xatt.SaveFile apath & nfname, false
  	  xsql = "INSERT INTO imageFile(newFileName, oldFileName) VALUES(" _
  	  	& pkStr(nfname,",") & pkStr(ofname,")")
  	  conn.execute xsql
	end if
  Next

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"

	conn.Execute(SQL)  
end sub '---- doUpdateDB() ----

sub showDoneBox()
	doneURI= ""
	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>sW{</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	alert("sWI")
	document.location.href = "<%=doneURI%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
