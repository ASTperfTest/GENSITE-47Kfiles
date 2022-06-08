<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="電子報管理"
HTProgFunc="新增發行"
HTUploadPath="/public/"
HTProgCode="GW1M51"
HTProgPrefix="ePub" %>
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
	dim xNewIdentity
 apath=server.mappath(HTUploadPath) & "\"
' response.write apath & "<HR>"
 if request.querystring("phase")<>"edit" then
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
	IF reg.htx_pubDate.value = Empty Then 
		MsgBox replace(nMsg,"{0}","發行日期"), 64, "Sorry!"
		reg.htx_pubDate.focus
		exit sub
	END IF
	IF (reg.htx_pubDate.value <> "") AND (NOT isDate(reg.htx_pubDate.value)) Then
		MsgBox replace(dMsg,"{0}","發行日期"), 64, "Sorry!"
		reg.htx_pubDate.focus
		exit sub
	END IF
	IF (reg.htx_ePubID.value <> "") AND (NOT isNumeric(reg.htx_ePubID.value)) Then
		MsgBox replace(iMsg,"{0}","發行ID"), 64, "Sorry!"
		reg.htx_ePubID.focus
		exit sub
	END IF
	IF reg.htx_title.value = Empty Then 
		MsgBox replace(nMsg,"{0}","標題"), 64, "Sorry!"
		reg.htx_title.focus
		exit sub
	END IF
	IF blen(reg.htx_title.value) > 80 Then
		MsgBox replace(replace(lMsg,"{0}","標題"),"{1}","80"), 64, "Sorry!"
		reg.htx_title.focus
		exit sub
	END IF
	IF reg.htx_dbDate.value = Empty Then 
		MsgBox replace(nMsg,"{0}","資料範圍起日"), 64, "Sorry!"
		reg.htx_dbDate.focus
		exit sub
	END IF
	IF (reg.htx_dbDate.value <> "") AND (NOT isDate(reg.htx_dbDate.value)) Then
		MsgBox replace(dMsg,"{0}","資料範圍起日"), 64, "Sorry!"
		reg.htx_dbDate.focus
		exit sub
	END IF
	IF reg.htx_deDate.value = Empty Then 
		MsgBox replace(nMsg,"{0}","資料範圍迄日"), 64, "Sorry!"
		reg.htx_deDate.focus
		exit sub
	END IF
	IF (reg.htx_deDate.value <> "") AND (NOT isDate(reg.htx_deDate.value)) Then
		MsgBox replace(dMsg,"{0}","資料範圍迄日"), 64, "Sorry!"
		reg.htx_deDate.focus
		exit sub
	END IF
	IF reg.htx_maxNo.value = Empty Then 
		MsgBox replace(nMsg,"{0}","資料則數"), 64, "Sorry!"
		reg.htx_maxNo.focus
		exit sub
	END IF
	IF (reg.htx_maxNo.value <> "") AND (NOT isNumeric(reg.htx_maxNo.value)) Then
		MsgBox replace(iMsg,"{0}","資料則數"), 64, "Sorry!"
		reg.htx_maxNo.focus
		exit sub
	END IF
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="ePubForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>--<%=session("ePaperName")%>】</td>
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

	sql = "INSERT INTO EpPub(ctRootId, "
	sqlValue = ") VALUES(" & session("epTreeID") & ","

	'2005/3/17 Apple 新增各個電子報自己的期別
	Set rsCheck = conn.execute("sp_columns @table_name = 'EpPub' , @column_name ='eIssue'")
	If Not rsCheck.EOF Then
		Set rsIssue = conn.execute( "select Max(""eIssue"") from EpPub where ctRootId='"& session("epTreeID") &"'" )
		If Not rsIssue.EOF Then
			sql = sql & "eIssue, "
			sqlValue = sqlValue  & rsIssue(0)+1 & ","
		End If 
	End If 
	
	IF xUpForm("htx_pubDate") <> "" Then
		sql = sql & "pubDate" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_pubDate"),",")
	END IF
	IF xUpForm("htx_ePubID") <> "" Then
		sql = sql & "ePubID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ePubID"),",")
	END IF
	IF xUpForm("htx_title") <> "" Then
		sql = sql & "title" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_title"),",")
	END IF
	IF xUpForm("htx_dbDate") <> "" Then
		sql = sql & "dbDate" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_dbDate"),",")
	END IF
	IF xUpForm("htx_deDate") <> "" Then
		sql = sql & "deDate" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_deDate"),",")
	END IF
	IF xUpForm("htx_maxNo") <> "" Then
		sql = sql & "maxNo" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_maxNo"),",")
	END IF
	IF xUpForm("htx_pubType") <> "" Then
		sql = sql & "pubType" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_pubType"),",")
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
  	  response.Write sql
  	  response.End	
  	  'conn.execute xsql
	end if
end if		
  Next

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
	set RSx = conn.Execute(SQL)
	xNewIdentity = RSx(0)	

end sub '---- doUpdateDB() ----

sub showDoneBox()
	doneURI= ""
	if doneURI = "" then	
		doneURI = HTprogPrefix & "List.asp?epTreeID=" & xUpForm("epTreeID")
	end if
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
	alert("新增完成！")
	document.location.href = "<%=doneURI%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
