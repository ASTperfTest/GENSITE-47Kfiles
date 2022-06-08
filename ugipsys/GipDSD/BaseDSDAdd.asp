<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap="單元資料定義"
HTProgFunc="新增"
HTUploadPath="/public/"
HTProgCode="GE1T01"
HTProgPrefix="BaseDsd" %>
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
'response.write apath & "<HR>"
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
	IF reg.htx_sbaseTableName.value<>"" and blen(reg.htx_sbaseTableName.value) > 30 Then
		MsgBox replace(replace(lMsg,"{0}","資料表名稱"),"{1}","40"), 64, "Sorry!"
		reg.htx_sbaseTableName.focus
		exit sub
	END IF	
	IF reg.htx_sbaseDsdname.value = Empty Then 
		MsgBox replace(nMsg,"{0}","資料單元名稱"), 64, "Sorry!"
		reg.htx_sbaseDsdname.focus
		exit sub
	END IF
	IF blen(reg.htx_sbaseDsdname.value) > 40 Then
		MsgBox replace(replace(lMsg,"{0}","資料單元名稱"),"{1}","40"), 64, "Sorry!"
		reg.htx_sbaseDsdname.focus
		exit sub
	END IF
	IF (reg.htx_ibaseDsd.value <> "") AND (NOT isNumeric(reg.htx_ibaseDsd.value)) Then
		MsgBox replace(iMsg,"{0}","單元資料定義ID"), 64, "Sorry!"
		reg.htx_ibaseDsd.focus
		exit sub
	END IF
	IF blen(reg.htx_tDesc.value) > 300 Then
		MsgBox replace(replace(lMsg,"{0}","說明"),"{1}","300"), 64, "Sorry!"
		reg.htx_tDesc.focus
		exit sub
	END IF
	IF reg.htx_inUse.value = Empty Then 
		MsgBox replace(nMsg,"{0}","是否生效"), 64, "Sorry!"
		reg.htx_inUse.focus
		exit sub
	END IF
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="BaseDsdForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
		<%if (HTProgRight and 1)=1 then%>
			<A href="BaseDsdQuery.asp" title="重設查詢條件">查詢</A>
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
	sql = "INSERT INTO BaseDsd("
	sqlValue = ") VALUES("
	IF xUpForm("htx_sbaseTableName") <> "" Then
		sql = sql & "sbaseTableName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_sbaseTableName"),",")
	END IF	
	IF xUpForm("htx_sbaseDsdname") <> "" Then
		sql = sql & "sbaseDsdname" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_sbaseDsdname"),",")
	END IF
	IF xUpForm("htx_ibaseDsd") <> "" Then
		sql = sql & "ibaseDsd" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ibaseDsd"),",")
	END IF
	IF xUpForm("htx_tDesc") <> "" Then
		sql = sql & "tDesc" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_tDesc"),",")
	END IF
	IF xUpForm("htx_rdsdcat") <> "" Then
		sql = sql & "rdsdcat" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_rdsdcat"),",")
	END IF
	IF xUpForm("htx_inUse") <> "" Then
		sql = sql & "inUse" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_inUse"),",")
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
	sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
	set RSx = conn.Execute(SQL)
	xNewIdentity = RSx(0)
	'----若無資料表名稱,則以CuDTx???為tablename
	IF xUpForm("htx_sbaseTableName") = "" Then
		xsbaseTableName="CuDTx"+cStr(xNewIdentity)
		sql = "Update BaseDsd set sbaseTableName='" &xsbaseTableName& "' where ibaseDsd="&xNewIdentity
		conn.execute(sql)
	END IF	
	'----若為多媒體型資料範本,自動新增Slave欄位
	IF xUpForm("htx_rDSDCat") = "MMO" Then
		SQLD=""
		SQLD=SQLD&"Insert Into BaseDSDField values("&xNewIdentity&",'mmofolderId',510,'Y','物件檔案存放目錄',null,'C',4,'N','N','N',null,'textbox',null,null,null,null,10);"
		SQLD=SQLD&"Insert Into BaseDSDField values("&xNewIdentity&",'mmoFileName',520,'Y','物件檔案名稱',null,'A',100,'N','N','N',null,'MMOfile',null,null,null,null,10);"
		SQLD=SQLD&"Insert Into BaseDSDField values("&xNewIdentity&",'mmoFileType',530,'Y','物件檔案類型',null,'A',50,'N','N','N',null,'refSelect',null,'MMOFileType',null,null,10);"
		SQLD=SQLD&"Insert Into BaseDSDField values("&xNewIdentity&",'mmoFileAlt',540,'Y','物件檔案替代文字',null,'A',50,'Y','N','N',null,'textbox',null,null,null,null,10);"
		SQLD=SQLD&"Insert Into BaseDSDField values("&xNewIdentity&",'mmoFileIcon',550,'Y','物件檔案顯示圖示',null,'A',50,'Y','N','N',null,'MMOfile',null,null,null,null,10);"  
		SQLD=SQLD&"Insert Into BaseDSDField values("&xNewIdentity&",'mmoFileWidth',560,'Y','物件檔案寬',null,'A',10,'Y','N','N',null,'textbox',null,null,null,null,10);"
		SQLD=SQLD&"Insert Into BaseDSDField values("&xNewIdentity&",'mmoFileHeight',570,'Y','物件檔案高',null,'A',10,'Y','N','N',null,'textbox',null,null,null,null,10);"  
		conn.execute(SQLD)
	END IF
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
