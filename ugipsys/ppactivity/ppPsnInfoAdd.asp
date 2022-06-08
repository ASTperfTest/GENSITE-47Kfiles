<% Response.Expires = 0
HTProgCap="活動學員管理"
HTProgFunc="新增"
HTUploadPath="/public/"
HTProgCode="PA010"
HTProgPrefix="ppPsnInfo" %>
<!--#Include file="../inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增表單</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
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

sub clientInitForm	'---- 新增時的表單預設值放在這裡
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
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
'msg1 = "請務必填寫「客戶編號」，不得為空白！"
'msg2 = "請務必填寫「客戶名稱」，不得為空白！"
'
'  If reg.pfx_ClientID.value = Empty Then
'     MsgBox msg1, 64, "Sorry!"
'     settab 0
'     reg.pfx_ClientID.focus
'     Exit Sub
'  End if
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
	IF reg.htx_myPassword.value = Empty Then 
		MsgBox replace(nMsg,"{0}","登入密碼"), 64, "Sorry!"
		reg.htx_myPassword.focus
		exit sub
	END IF
	IF reg.htx_myPassword.value <>  reg.htx_myPassword2check.value  Then 
		MsgBox "密碼欄與確認欄的輸入資料需相同", 64, "Sorry!"
		reg.htx_myPassword.focus
		exit sub
	END IF
	IF reg.htx_psnID.value = Empty Then 
		MsgBox replace(nMsg,"{0}","身份證號"), 64, "Sorry!"
		reg.htx_psnID.focus
		exit sub
	END IF
	IF blen(reg.htx_psnID.value) > 10 Then
		MsgBox replace(replace(lMsg,"{0}","身份證號"),"{1}","10"), 64, "Sorry!"
		reg.htx_psnID.focus
		exit sub
	END IF
	IF reg.htx_pName.value = Empty Then 
		MsgBox replace(nMsg,"{0}","姓名"), 64, "Sorry!"
		reg.htx_pName.focus
		exit sub
	END IF
	IF blen(reg.htx_pName.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","姓名"),"{1}","20"), 64, "Sorry!"
		reg.htx_pName.focus
		exit sub
	END IF
	IF reg.htx_birthDay.value = Empty Then 
		MsgBox replace(nMsg,"{0}","出生日"), 64, "Sorry!"
		reg.htx_birthDay.focus
		exit sub
	END IF
	IF (reg.htx_birthDay.value <> "") AND (NOT isDate(reg.htx_birthDay.value)) Then
		MsgBox replace(dMsg,"{0}","出生日"), 64, "Sorry!"
		reg.htx_birthDay.focus
		exit sub
	END IF
	IF blen(reg.htx_eMail.value) > 50 Then
		MsgBox replace(replace(lMsg,"{0}","eMail"),"{1}","50"), 64, "Sorry!"
		reg.htx_eMail.focus
		exit sub
	END IF
	IF reg.htx_tel.value = Empty Then 
		MsgBox replace(nMsg,"{0}","連絡電話"), 64, "Sorry!"
		reg.htx_tel.focus
		exit sub
	END IF
	IF blen(reg.htx_tel.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","連絡電話"),"{1}","20"), 64, "Sorry!"
		reg.htx_tel.focus
		exit sub
	END IF
	IF reg.htx_emergContact.value = Empty Then 
		MsgBox replace(nMsg,"{0}","聯絡地址"), 64, "Sorry!"
		reg.htx_emergContact.focus
		exit sub
	END IF
	IF blen(reg.htx_emergContact.value) > 60 Then
		MsgBox replace(replace(lMsg,"{0}","緊急聯絡人姓名及電話"),"{1}","60"), 64, "Sorry!"
		reg.htx_emergContact.focus
		exit sub
	END IF
	IF blen(reg.htx_deptName.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","工作部門"),"{1}","20"), 64, "Sorry!"
		reg.htx_deptName.focus
		exit sub
	END IF
	IF blen(reg.htx_jobName.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","職稱"),"{1}","20"), 64, "Sorry!"
		reg.htx_jobName.focus
		exit sub
	END IF
	IF blen(reg.htx_corpName.value) > 60 Then
		MsgBox replace(replace(lMsg,"{0}","公司全銜"),"{1}","60"), 64, "Sorry!"
		reg.htx_corpName.focus
		exit sub
	END IF
	IF blen(reg.htx_corpID.value) > 8 Then
		MsgBox replace(replace(lMsg,"{0}","統一編號"),"{1}","8"), 64, "Sorry!"
		reg.htx_corpID.focus
		exit sub
	END IF
	IF blen(reg.htx_corpZip.value) > 3 Then
		MsgBox replace(replace(lMsg,"{0}","郵遞區號"),"{1}","3"), 64, "Sorry!"
		reg.htx_corpZip.focus
		exit sub
	END IF
	IF blen(reg.htx_corpAddr.value) > 60 Then
		MsgBox replace(replace(lMsg,"{0}","聯絡地址"),"{1}","60"), 64, "Sorry!"
		reg.htx_corpAddr.focus
		exit sub
	END IF
	IF blen(reg.htx_corpContact.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","聯絡人"),"{1}","20"), 64, "Sorry!"
		reg.htx_corpContact.focus
		exit sub
	END IF
	IF blen(reg.htx_corpTel.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","電話"),"{1}","20"), 64, "Sorry!"
		reg.htx_corpTel.focus
		exit sub
	END IF
	IF blen(reg.htx_corpFax.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","傳真"),"{1}","20"), 64, "Sorry!"
		reg.htx_corpFax.focus
		exit sub
	END IF
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="ppPsnInfoForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
		<%if (HTProgRight and 1)=1 then%>
			<A href="ppPsnInfoQuery.asp" title="重設查詢條件">查詢</A>
		<%end if%>
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A> 
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

'---- 後端資料驗證檢查程式碼放在這裡	，如下例，有錯時設 errMsg="xxx" 並 exit sub ------
'	SQL = "Select * From Client Where ClientID = '"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "「客戶編號」重複!!請重新建立客戶編號!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

sub doUpdateDB()
	sql = "INSERT INTO paPsnInfo("
	sqlValue = ") VALUES("
	IF xUpForm("htx_myPassword") <> "" Then
		sql = sql & "myPassword" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_myPassword"),",")
	END IF
	IF xUpForm("htx_psnID") <> "" Then
		sql = sql & "psnID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_psnID"),",")
	END IF
	IF xUpForm("htx_pName") <> "" Then
		sql = sql & "pName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_pName"),",")
	END IF
	IF xUpForm("htx_birthDay") <> "" Then
		sql = sql & "birthDay" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_birthDay"),",")
	END IF
	IF xUpForm("htx_eMail") <> "" Then
		sql = sql & "eMail" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_eMail"),",")
	END IF
	IF xUpForm("htx_tel") <> "" Then
		sql = sql & "tel" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_tel"),",")
	END IF
	IF xUpForm("htx_emergContact") <> "" Then
		sql = sql & "emergContact" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_emergContact"),",")
	END IF
	IF xUpForm("htx_deptName") <> "" Then
		sql = sql & "deptName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_deptName"),",")
	END IF
	IF xUpForm("htx_jobName") <> "" Then
		sql = sql & "jobName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_jobName"),",")
	END IF
	IF xUpForm("htx_topEdu") <> "" Then
		sql = sql & "topEdu" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_topEdu"),",")
	END IF
	IF xUpForm("htx_meatKind") <> "" Then
		sql = sql & "meatKind" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_meatKind"),",")
	END IF
	IF xUpForm("htx_corpName") <> "" Then
		sql = sql & "corpName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_corpName"),",")
	END IF
	IF xUpForm("htx_corpID") <> "" Then
		sql = sql & "corpID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_corpID"),",")
	END IF
	IF xUpForm("htx_corpZip") <> "" Then
		sql = sql & "corpZip" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_corpZip"),",")
	END IF
	IF xUpForm("htx_corpAddr") <> "" Then
		sql = sql & "corpAddr" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_corpAddr"),",")
	END IF
	IF xUpForm("htx_corpContact") <> "" Then
		sql = sql & "corpContact" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_corpContact"),",")
	END IF
	IF xUpForm("htx_corpTel") <> "" Then
		sql = sql & "corpTel" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_corpTel"),",")
	END IF
	IF xUpForm("htx_corpFax") <> "" Then
		sql = sql & "corpFax" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_corpFax"),",")
	END IF
	IF xUpForm("htx_datafrom") <> "" Then
		sql = sql & "datafrom" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_datafrom"),",")
	END IF
	IF xUpForm("htx_mobile") <> "" Then
		sql = sql & "mobile" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_mobile"),",")
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
	conn.Execute SQL

end sub '---- doUpdateDB() ----

sub showDoneBox()
	doneURI= ""
	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	alert("新增完成！")
	document.location.href = "<%=doneURI%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
