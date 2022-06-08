﻿<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap="單元資料定義"
HTProgFunc="新增欄位"
HTUploadPath="/public/"
HTProgCode="GE1T01"
HTProgPrefix="bDSDField" %>
<!--#include virtual = "/inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增表單</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
dim xNewIdentity,xSeq
 apath=server.mappath(HTUploadPath) & "\"
' response.write apath & "<HR>"
 if request.querystring("phase")<>"add" then
Set xup = Server.CreateObject("TABS.Upload")
xup.codepage=65001
xup.Start apath
else
Set xup = Server.CreateObject("TABS.Upload")
end if

function xUpForm(xvar)
	xUpForm = xup.form(xvar)
end function

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
	sqlCom = "SELECT max(xfieldSeq) FROM BaseDsdfield AS htx WHERE htx.ibaseDsd=" & pkStr(request.queryString("ibaseDsd"),"")
	Set RSreg = Conn.execute(sqlcom)
	if isNull(RSreg(0)) then
		xSeq=510
	else
		xSeq=RSreg(0)+10	
	end if
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

sub clientInitForm	'---- 新增時的表單預設值放在這裡
	reg.htx_ibaseDsd.value= "<%=request("ibaseDsd")%>"
	reg.htx_xfieldSeq.value= "<%=xSeq%>"	
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
	IF (reg.htx_ibaseDsd.value <> "") AND (NOT isNumeric(reg.htx_ibaseDsd.value)) Then
		MsgBox replace(iMsg,"{0}","單元資料定義ID"), 64, "Sorry!"
		reg.htx_ibaseDsd.focus
		exit sub
	END IF
	IF (reg.htx_iBaseField.value <> "") AND (NOT isNumeric(reg.htx_iBaseField.value)) Then
		MsgBox replace(iMsg,"{0}","單元資料欄位ID"), 64, "Sorry!"
		reg.htx_iBaseField.focus
		exit sub
	END IF
	IF reg.htx_inUse.value = Empty Then 
		MsgBox replace(nMsg,"{0}","是否生效"), 64, "Sorry!"
		reg.htx_inUse.focus
		exit sub
	END IF
	IF blen(reg.htx_xfieldLabel.value) > 30 Then
		MsgBox replace(replace(lMsg,"{0}","標題"),"{1}","30"), 64, "Sorry!"
		reg.htx_xfieldLabel.focus
		exit sub
	END IF
	IF blen(reg.htx_xfieldDesc.value) > 50 Then
		MsgBox replace(replace(lMsg,"{0}","說明"),"{1}","50"), 64, "Sorry!"
		reg.htx_xfieldDesc.focus
		exit sub
	END IF
	IF reg.htx_xdataType.value = Empty Then 
		MsgBox replace(nMsg,"{0}","資料型別"), 64, "Sorry!"
		reg.htx_xdataType.focus
		exit sub
	END IF
	IF (reg.htx_xdataLen.value <> "") AND (NOT isNumeric(reg.htx_xdataLen.value)) Then
		MsgBox replace(iMsg,"{0}","資料長度"), 64, "Sorry!"
		reg.htx_xdataLen.focus
		exit sub
	END IF
	IF reg.htx_xinputType.value = Empty Then 
		MsgBox replace(nMsg,"{0}","輸入型式"), 64, "Sorry!"
		reg.htx_xinputType.focus
		exit sub
	END IF
	IF blen(reg.htx_xrefLookup.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","xrefLookup"),"{1}","20"), 64, "Sorry!"
		reg.htx_xrefLookup.focus
		exit sub
	END IF
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="bDSDFieldForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
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
	sql = "INSERT INTO BaseDsdfield("
	sqlValue = ") VALUES("
	IF xUpForm("htx_xfieldSeq") <> "" Then
		sql = sql & "xfieldSeq" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xfieldSeq"),",")
	END IF
	IF xUpForm("htx_ibaseDsd") <> "" Then
		sql = sql & "ibaseDsd" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ibaseDsd"),",")
	END IF
	IF xUpForm("htx_iBaseField") <> "" Then
		sql = sql & "iBaseField" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_iBaseField"),",")
	END IF
	IF xUpForm("htx_inUse") <> "" Then
		sql = sql & "inUse" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_inUse"),",")
	END IF
	IF xUpForm("htx_xfieldLabel") <> "" Then
		sql = sql & "xfieldLabel" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xfieldLabel"),",")
	END IF
	IF xUpForm("htx_xfieldDesc") <> "" Then
		sql = sql & "xfieldDesc" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xfieldDesc"),",")
	END IF
	IF xUpForm("htx_xdataType") <> "" Then
		sql = sql & "xdataType" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xdataType"),",")
	END IF
	IF xUpForm("htx_xdataLen") <> "" Then
		sql = sql & "xdataLen" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xdataLen"),",")
	END IF
	IF xUpForm("htx_xcanNull") <> "" Then
		sql = sql & "xcanNull" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xcanNull"),",")
	END IF
	IF xUpForm("htx_xinputType") <> "" Then
		sql = sql & "xinputType" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xinputType"),",")
	END IF
	IF xUpForm("htx_xrefLookup") <> "" Then
		sql = sql & "xrefLookup" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xrefLookup"),",")
	END IF
	IF xUpForm("htx_xfieldName") <> "" Then
		sql = sql & "xfieldName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xfieldName"),",")
	END IF
	
 For Each Form In xup.Form
 If Form.IsFile Then 
    if left(Form.Name,6) = "htImg_" then
	  ofname = Form.FileName
	  fnExt = ""
	  if instr(ofname, ".")>0 then	fnext=mid(ofname, instr(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(Form.Name,7) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xup.Form(Form.Name).SaveAs apath & nfname, True
  	elseif left(Form.Name,7) = "htFile_" then
	  ofname = Form.FileName
	  fnExt = ""
	  if instr(ofname, ".")>0 then	fnext=mid(ofname, instr(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(Form.Name,8) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xup.Form(Form.Name).SaveAs apath & nfname, True
  	  xsql = "INSERT INTO imageFile(newFileName, oldFileName) VALUES(" _
  	  	& pkStr(nfname,",") & pkStr(ofname,")")
  	  conn.execute xsql
	end if
end if		
  Next

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"

	conn.Execute(SQL)  
end sub '---- doUpdateDB() ----

sub showDoneBox()
	doneURI= "BaseDSDEditList.asp"
	doneURI= doneURI & "?phase=edit&ibaseDsd=" & xUpForm("htx_ibaseDsd")
	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp?" & "phase=edit"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	alert("<%=doneURI%>新增完成！")
	document.location.href = "<%=doneURI%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
