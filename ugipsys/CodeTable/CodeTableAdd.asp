<%@ CodePage = 65001 %>
<%HTProgCap="NXwq"
HTProgCode="Pn50M03"
HTProgPrefix="CodeTable" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
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
	ClientID=0
	taskLable="sW" & HTProgCap

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

sub clientInitForm	
	reg.tfx_CodeTblName.value="CodeMain"
	reg.tfx_CodeSrcFld.value="codeMetaID"	
	reg.tfx_CodeValueFld.value="mCode"	
	reg.tfx_CodeValueFldName.value="代碼欄"
	reg.tfx_CodeDisplayFld.value="mValue"
	reg.tfx_CodeDisplayFldName.value="顯示欄"
	reg.tfx_CodeSortFld.value="mSortValue"
	reg.tfx_CodeSortFldName.value="排序欄"	
	reg.tfx_CodeType.value=""				
	reg.bfx_ShowOrNot.checked=true
end sub

sub window_onLoad
	clientInitForm
end sub
</script>	
<% end sub '---- initForm() ----%>

<% sub showForm() 	'===================== Client Side Validation Put HERE =========== %>
<script language="vbscript">

Sub formAddSubmit()	
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
msg1 = "請務必輸入「代碼ID」，不得為空白！"
'msg1 = "Please insert [Code ID]!"
msg2 = "請務必輸入「代碼名稱」，不得為空白！"
'msg2 = "Please insert [Code Name]!"
msg3 = "請務必輸入「存放資料表」，不得為空白！"
'msg3 = "Please insert [Code Table]!"
msg4 = "請務必輸入「Value值欄」，不得為空白！"
'msg4 = "Please insert [Value]!"
msg5 = "請務必輸入「Display欄」，不得為空白！"
'msg5 = "Please insert [Display]!"
msg6 = "請務必輸入「過濾(排序)欄」，不得為空白！"
'msg6 = "Please insert [Sort]!"

  If reg.pfx_CodeID.value = Empty Then
     MsgBox msg1, 64, "Sorry!"
     reg.pfx_CodeID.focus
     Exit Sub
  End if
  If reg.tfx_CodeName.value = Empty Then
     MsgBox msg2, 64, "Sorry!"
     reg.tfx_CodeName.focus
     Exit Sub
  End if
  If reg.tfx_CodeTblName.value = Empty Then
     MsgBox msg3, 64, "Sorry!"
     reg.tfx_CodeTblName.focus
     Exit Sub
  End if
  If reg.tfx_CodeValueFld.value = Empty Then
     MsgBox msg4, 64, "Sorry!"
     reg.tfx_CodeValueFld.focus
     Exit Sub
  End if
  If reg.tfx_CodeDisplayFld.value = Empty Then
     MsgBox msg5, 64, "Sorry!"
     reg.tfx_CodeDisplayFld.focus
     Exit Sub
  End if
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="CodeTableForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增表單</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2"><%=Title%> </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" width="60%">【新增代碼定義】標有<font color=red>※</font>符號欄位為必填欄位</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><% if (HTProgRight and 2)=2 then %><a href="<%=HTprogPrefix%>Query.asp">查詢</a><% End IF %>&nbsp;</td>
  </tr>  
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>  
  <tr>
    <td width="80%" height=230 valign=top colspan="2">
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
	SQL = "Select codeId From CodeMetaDef Where codeId = N'"& Request("pfx_codeId") &"'"
	set RSvalidate = conn.Execute(SQL)

	if not RSvalidate.EOF Then
		errMsg = "uNXIDv!!Эsإ!"
		exit sub
	end if

end sub '---- checkDBValid() ----

sub doUpdateDB()
	xShowOrNot="N"
	if request("bfx_ShowOrNot")<>"" then xShowOrNot="Y"
	sql = "INSERT INTO CodeMetaDef(showOrNot,"
	sqlValue = ") VALUES(N'"&xshowOrNot&"',"
	if checkGIPconfig("codeML") then
		xisML="N"
		if request("bfx_isML")<>"" then xisML="Y"
		sql = sql & "isML,"
		sqlValue = sqlValue & pkStr(xisML,",")
	end if
	for each x in request.form
	 if request(x) <> "" then
	  if mid(x,2,3) = "fx_" AND left(x,1)<>"b" then
		select case left(x,1)
		  case "p"
			sql = sql & mid(x,5) & ","
			sqlValue = sqlValue & pkStr(request(x),",")
		  case "d"
			sql = sql & mid(x,5) & ","
			sqlValue = sqlValue & pkStr(request(x),",")
		  case "n"
			sql = sql & mid(x,5) & ","
			sqlValue = sqlValue & request(x) & ","
		  case else
			sql = sql & mid(x,5) & ","
			sqlValue = sqlValue & pkStr(request(x),",")
		end select
	  end if
	 end if
	next 
	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"

	conn.Execute(SQL)  
end sub '---- doUpdateDB() ----

%>

<% sub showDoneBox() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../setstyle.css">
</head>
<body>
<script language=vbs>
	alert("Done!")
	document.location.href = "<%=HTprogPrefix%>Query.asp"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
