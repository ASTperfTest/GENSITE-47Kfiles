<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap="主題單元管理"
HTProgFunc="新增"
HTUploadPath="/public/"
HTProgCode="GE1T11"
HTProgPrefix="CtUnit" %>
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
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
	dim xNewIdentity
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
	reg.htx_checkYn.value="N"
	reg.htx_newWindow.value= "N"
	reg.htx_fctUnitOnly.value= "Y"
	<%if checkGIPconfig("MMOFolder") then %>
		document.all("MMOFolder").style.display="none"
<%	end if %>		
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
xrefDSDCat=""        
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
	IF blen(reg.htx_ctUnitName.value) > 200 Then
		MsgBox replace(replace(lMsg,"{0}","主題單元名稱"),"{1}","200"), 64, "Sorry!"
		reg.htx_ctUnitName.focus
		exit sub
	END IF
	IF (reg.htx_CtUnitID.value <> "") AND (NOT isNumeric(reg.htx_CtUnitID.value)) Then
		MsgBox replace(iMsg,"{0}","主題單元ID"), 64, "Sorry!"
		reg.htx_CtUnitID.focus
		exit sub
	END IF
	IF reg.htImg_CtUnitLogo.value <> "" Then
		xIMGname = reg.htImg_CtUnitLogo.value
		xFileType = ""
		if instr(xIMGname, ".")>0 then	xFileType=lcase(mid(xIMGname, instr(xIMGname, ".")+1))
		IF xFileType<>"gif" AND xFileType<>"jpg" AND xFileType<>"jpeg" then
			MsgBox replace(pMsg,"{0}","單元LOGO"), 64, "Sorry!"
			reg.htImg_CtUnitLogo.focus
			exit sub
		END IF
	END IF
	IF reg.htx_ctUnitKind.value = Empty Then 
		MsgBox replace(nMsg,"{0}","單元類型"), 64, "Sorry!"
		reg.htx_ctUnitKind.focus
		exit sub
	END IF
	IF blen(reg.htx_redirectURL.value) > 100 Then
		MsgBox replace(replace(lMsg,"{0}","重導URL"),"{1}","100"), 64, "Sorry!"
		reg.htx_redirectURL.focus
		exit sub
	END IF
<% 	if checkGIPconfig("CtUnitExpireCheck") then %>
	IF (reg.htx_CtUnitExpireDay.value <> "") AND (NOT isNumeric(reg.htx_CtUnitExpireDay.value)) Then
		MsgBox replace(iMsg,"{0}","單元資料逾期天數"), 64, "Sorry!"
		reg.htx_CtUnitExpireDay.focus
		exit sub
	END IF
<%	end if %>
	IF (reg.htx_ibaseDsd.value <> "") AND (NOT isNumeric(reg.htx_ibaseDsd.value)) Then
		MsgBox replace(iMsg,"{0}","單元資料定義"), 64, "Sorry!"
		reg.htx_ibaseDsd.focus
		exit sub
	END IF
	IF reg.htx_fctUnitOnly.value = Empty Then 
		MsgBox replace(nMsg,"{0}","只顯示此單元資料"), 64, "Sorry!"
		reg.htx_fctUnitOnly.focus
		exit sub
	END IF
	<%if checkGIPconfig("MMOFolder") then %>
	    if reg.xrefDSDCat.value="MMO" and reg.htx_MMOFolderID.value="0" then
		alert "請點選多媒體型主題單元存放根目錄!"
		reg.htx_MMOFolderID.focus
		exit sub		
	    end if	
	<%end if%>  
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub
sub iBaseDSDSpilt()
	if reg.iBaseDSDSelect.value="" then 
		document.all("MMOFolder").style.display="none"
		reg.htx_iBaseDSD.value=""
		reg.xrefDSDCat.value=""
		<%if checkGIPconfig("MMOFolder") then %>
		        reg.htx_MMOFolderID.value="0"
		<%end if%>
		exit sub
	end if
	BaseDSDPos=instr(reg.iBaseDSDSelect.value,"---")
	reg.htx_iBaseDSD.value=Left(reg.iBaseDSDSelect.value,BaseDSDPos-1)
	reg.xrefDSDCat.value=mid(reg.iBaseDSDSelect.value,BaseDSDPos+3)
	<%if checkGIPconfig("MMOFolder") then %>
	    if reg.xrefDSDCat.value="MMO" then
		document.all("MMOFolder").style.display=""
	    else
	        document.all("MMOFolder").style.display="none"
	        reg.htx_MMOFolderID.value="0"
	    end if
	<%end if%>
end sub
<%if checkGIPconfig("MMOFolder") then %>
sub MMOFolderAdd()
	if reg.htx_MMOFolderID.value="0" then
		alert "請先選擇存放根目錄!"
		reg.htx_MMOFolderID.focus
		exit sub				
	else
		window.open "/MMO/MMOAddFolder.asp?S=11&MMOFolderID="&reg.htx_MMOFolderID.value,"","height=400,width=550"
	end if
end sub
sub MMOFolderAddReload()
	document.location.reload()
end sub
<%end if%>
</script>

<!--#include file="CtUnitForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
		<%if (HTProgRight and 1)=1 then%>
			<A href="CtUnitQuery.asp" title="重設查詢條件">查詢</A>
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
	sql = "INSERT INTO CtUnit("
	sqlValue = ") VALUES("
	IF xUpForm("htx_ctUnitName") <> "" Then
		sql = sql & "ctUnitName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ctUnitName"),",")
	END IF
	IF xUpForm("htx_CtUnitID") <> "" Then
		sql = sql & "CtUnitID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_CtUnitID"),",")
	END IF
	IF xUpForm("htx_ctUnitKind") <> "" Then
		sql = sql & "ctUnitKind" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ctUnitKind"),",")
	END IF
	IF xUpForm("htx_redirectURL") <> "" Then
		sql = sql & "redirectURL" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_redirectURL"),",")
	END IF
	IF xUpForm("htx_newWindow") <> "" Then
		sql = sql & "newWindow" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_newWindow"),",")
	END IF
	IF xUpForm("htx_ibaseDsd") <> "" Then
		sql = sql & "ibaseDsd" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ibaseDsd"),",")
	END IF
	IF xUpForm("htx_fctUnitOnly") <> "" Then
		sql = sql & "fctUnitOnly" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_fctUnitOnly"),",")
	END IF
	IF xUpForm("htx_checkYn") <> "" Then
		sql = sql & "checkYn" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_checkYn"),",")
	END IF
	IF xUpForm("htx_deptID") <> "" Then
		sql = sql & "deptID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_deptID"),",")
	END IF
 	if checkGIPconfig("CtUnitExpireCheck") then
		IF xUpForm("htx_CtUnitExpireDay") <> "" Then
			sql = sql & "CtUnitExpireDay" & ","
			sqlValue = sqlValue & pkStr(xUpForm("htx_CtUnitExpireDay"),",")
		END IF
	end if
	if checkGIPconfig("MMOFolder") then 
		IF xUpForm("htx_MMOFolderID") <> "" and xUpForm("htx_MMOFolderID") <> "0" Then
			sql = sql & "MMOFolderID" & ","
			sqlValue = sqlValue & pkStr(xUpForm("htx_MMOFolderID"),",")
		END IF
	end if

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
