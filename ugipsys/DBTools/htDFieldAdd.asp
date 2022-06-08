<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="資料物件欄位管理"
HTProgFunc="新增資料物件欄位"
HTProgCode="HT011"
HTProgPrefix="HtDfield" %>
<!--#include virtual = "/inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增表單</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
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

sub clientInitForm	'---- 新增時的表單預設值放在這裡
	reg.htx_entityId.value= "<%=request("entityId")%>"
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
	IF reg.htx_xfieldName.value = Empty Then
		MsgBox replace(nMsg,"{0}","資料欄位名稱"), 64, "Sorry!"
		reg.htx_xfieldName.focus
		exit sub
	END IF
	IF blen(reg.htx_xfieldName.value) > 30 Then
		MsgBox replace(replace(lMsg,"{0}","資料欄位名稱"),"{1}","30"), 64, "Sorry!"
		reg.htx_xfieldName.focus
		exit sub
	END IF
	IF reg.htx_entityId.value = Empty Then
		MsgBox replace(nMsg,"{0}","資料物件代碼"), 64, "Sorry!"
		reg.htx_entityId.focus
		exit sub
	END IF
	IF (reg.htx_xfieldSeq.value <> "") AND (NOT isNumeric(reg.htx_xfieldSeq.value)) Then
		MsgBox replace(iMsg,"{0}","排序值"), 64, "Sorry!"
		reg.htx_xfieldSeq.focus
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
	IF (reg.htx_xinputLen.value <> "") AND (NOT isNumeric(reg.htx_xinputLen.value)) Then
		MsgBox replace(iMsg,"{0}","輸入長度"), 64, "Sorry!"
		reg.htx_xinputLen.focus
		exit sub
	END IF				
	IF blen(reg.htx_xdefaultvalue.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","dbDefault"),"{1}","20"), 64, "Sorry!"
		reg.htx_xdefaultvalue.focus
		exit sub
	END IF
	IF blen(reg.htx_xclientDefault.value) > 100 Then
		MsgBox replace(replace(lMsg,"{0}","clientDefault"),"{1}","100"), 64, "Sorry!"
		reg.htx_xclientDefault.focus
		exit sub
	END IF
	IF blen(reg.htx_xrefLookup.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","xrefLookup"),"{1}","100"), 64, "Sorry!"
		reg.htx_xrefLookup.focus
		exit sub
	END IF	
	IF (reg.htx_xrows.value <> "") AND (NOT isNumeric(reg.htx_xrows.value)) Then
		MsgBox replace(iMsg,"{0}","rows"), 64, "Sorry!"
		reg.htx_xrows.focus
		exit sub
	END IF	  
	IF (reg.htx_xcols.value <> "") AND (NOT isNumeric(reg.htx_xcols.value)) Then
		MsgBox replace(iMsg,"{0}","cols"), 64, "Sorry!"
		reg.htx_xcols.focus
		exit sub
	END IF	  	  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="HtDfieldForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <a href="Javascript:window.history.back();">回前頁</a>
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
	sql = "INSERT INTO HtDfield("
	sqlValue = ") VALUES("
	IF request("htx_xfieldName") <> "" Then
		sql = sql & "xfieldName" & ","
		sqlValue = sqlValue & pkStr(request("htx_xfieldName"),",")
	END IF
	IF request("htx_entityId") <> "" Then
		sql = sql & "entityId" & ","
		sqlValue = sqlValue & pkStr(request("htx_entityId"),",")
	END IF
	IF request("htx_xfieldLabel") <> "" Then
		sql = sql & "xfieldLabel" & ","
		sqlValue = sqlValue & pkStr(request("htx_xfieldLabel"),",")
	END IF
	IF request("htx_xfieldDesc") <> "" Then
		sql = sql & "xfieldDesc" & ","
		sqlValue = sqlValue & pkStr(request("htx_xfieldDesc"),",")
	END IF
	IF request("htx_xfieldSeq") <> "" Then
		sql = sql & "xfieldSeq" & ","
		sqlValue = sqlValue & drn("htx_xfieldSeq")
	END IF		
	IF request("htx_xdataType") <> "" Then
		sql = sql & "xdataType" & ","
		sqlValue = sqlValue & pkStr(request("htx_xdataType"),",")
	END IF
	IF request("htx_xdataLen") <> "" Then
		sql = sql & "xdataLen" & ","
		sqlValue = sqlValue & drn("htx_xdataLen")
	END IF
	IF request("htx_xinputLen") <> "" Then
		sql = sql & "xinputLen" & ","
		sqlValue = sqlValue & drn("htx_xinputLen")
	END IF	
	IF request("htx_xcanNull") <> "" Then
		sql = sql & "xcanNull" & ","
		sqlValue = sqlValue & pkStr(request("htx_xcanNull"),",")
	END IF
	IF request("htx_xisPrimaryKey") <> "" Then
		sql = sql & "xisPrimaryKey" & ","
		sqlValue = sqlValue & pkStr(request("htx_xisPrimaryKey"),",")
	END IF
	IF request("htx_xidentity") <> "" Then
		sql = sql & "xidentity" & ","
		sqlValue = sqlValue & pkStr(request("htx_xidentity"),",")
	END IF
	IF request("htx_xdefaultvalue") <> "" Then
		sql = sql & "xdefaultvalue" & ","
		sqlValue = sqlValue & pkStr(request("htx_xdefaultvalue"),",")
	END IF
	IF request("htx_xclientDefault") <> "" Then
		sql = sql & "xclientDefault" & ","
		sqlValue = sqlValue & pkStr(request("htx_xclientDefault"),",")
	END IF
	IF request("htx_xinputType") <> "" Then
		sql = sql & "xinputType" & ","
		sqlValue = sqlValue & pkStr(request("htx_xinputType"),",")
	END IF
	IF request("htx_xrefLookup") <> "" Then
		sql = sql & "xrefLookup" & ","
		sqlValue = sqlValue & pkStr(request("htx_xrefLookup"),",")
	END IF		
	IF request("htx_xrows") <> "" Then
		sql = sql & "xrows" & ","
		sqlValue = sqlValue & drn("htx_xrows")
	END IF	
	IF request("htx_xcols") <> "" Then
		sql = sql & "xcols" & ","
		sqlValue = sqlValue & drn("htx_xcols")
	END IF		
	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"

	'response.write sql & "<HR>"
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
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	alert("新增完成！")
	document.location.href = "<%=HTprogPrefix%>List.asp?<%=request.serverVariables("QUERY_STRING")%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
