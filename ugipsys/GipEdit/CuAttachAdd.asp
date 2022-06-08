<%@ CodePage = 65001 %>
<% Response.Expires = 0
Server.ScriptTimeOut = 3000
HTProgCap="資料附件"
HTProgFunc="新增"
HTUploadPath=session("Public")+"Attachment/"
HTProgCode="GC1AP1"
HTProgPrefix="CuAttach" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
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
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function
dim xNewIdentity
'----ftp捊Bz
	FTPErrorMSG=""
	FTPfilePath="public/Attachment"
	SQLP = "Select * from UpLoadSite where upLoadSiteId='file'"
	Set RSP = conn.execute(SQLP)
	if not RSP.EOF  then
   		xFTPIP = RSP("UpLoadSiteFTPIP")
   		xFTPPort = RSP("UpLoadSiteFTPPort")
   		xFTPID = RSP("UpLoadSiteFTPID")
   		xFTPPWD = RSP("UpLoadSiteFTPPWD")
   	end if
'----ftp捊Bzend 
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

sub clientInitForm	'---- sW阞w]bo
	reg.htx_xiCuItem.value= "<%=request("iCuItem")%>"
	reg.htx_aeditor.value= "<%=session("userID")%>"
	reg.htx_aeditDate.value= "<%=date()%>"
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
	IF reg.htx_atitle.value = Empty Then 
		MsgBox replace(nMsg,"{0}","附件名稱"), 64, "Sorry!"
		reg.htx_atitle.focus
		exit sub
	END IF
	IF blen(reg.htx_atitle.value) > 80 Then
		MsgBox replace(replace(lMsg,"{0}","附件名稱"),"{1}","80"), 64, "Sorry!"
		reg.htx_atitle.focus
		exit sub
	END IF
	IF (reg.htx_xiCuItem.value <> "") AND (NOT isNumeric(reg.htx_xiCuItem.value)) Then
		MsgBox replace(iMsg,"{0}","資料ID"), 64, "Sorry!"
		reg.htx_xiCuItem.focus
		exit sub
	END IF
	IF (reg.htx_aeditDate.value <> "") AND (NOT isDate(reg.htx_aeditDate.value)) Then
		MsgBox replace(dMsg,"{0}","上傳日"), 64, "Sorry!"
		reg.htx_aeditDate.focus
		exit sub
	END IF
	IF (reg.htx_ixCuAttach.value <> "") AND (NOT isNumeric(reg.htx_ixCuAttach.value)) Then
		MsgBox replace(iMsg,"{0}","內部ID"), 64, "Sorry!"
		reg.htx_ixCuAttach.focus
		exit sub
	END IF
	IF blen(reg.htx_adesc.value) > 800 Then
		MsgBox replace(replace(lMsg,"{0}","說明"),"{1}","800"), 64, "Sorry!"
		reg.htx_adesc.focus
		exit sub
	END IF
	<%if checkGIPconfig("AttachLarge") then%>
	  	if reg.htFile_NFileType(0).checked = true then
		  	IF reg.htfile_nfileName.value = Empty Then 
				MsgBox replace(nMsg,"{0}","系統檔名"), 64, "Sorry!"
				reg.htfile_nfileName.focus
				exit sub
			END IF
	  	elseif reg.htFile_NFileType(1).checked = true then
		  	IF reg.htFile_nfileName_Large.value = Empty Then 
				MsgBox replace(nMsg,"{0}","系統檔名"), 64, "Sorry!"
				reg.htFile_nfileName_Large.focus
				exit sub
			END IF
	  	end if
	<%else%>
'		IF reg.htfile_nfileName.value = Empty Then 
'			MsgBox replace(nMsg,"{0}","系統檔名"), 64, "Sorry!"
'			reg.htfile_nfileName.focus
'			exit sub
'		END IF
	
	<%end if%>  
	reg.submitTask.value = "ADD"
	reg.Submit
End Sub

</script>


<!--#include file="CuAttachForm.inc"-->
                   
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
	sql = "INSERT INTO CuDtattach("
	sqlValue = ") VALUES("
	IF xUpForm("htx_atitle") <> "" Then
		sql = sql & "atitle" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_atitle"),",")
	END IF
	IF xUpForm("htx_xiCuItem") <> "" Then
		sql = sql & "xiCuItem" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xiCuItem"),",")
	END IF
	IF xUpForm("htx_aeditor") <> "" Then
		sql = sql & "aeditor" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_aeditor"),",")
	END IF
	IF xUpForm("htx_aeditDate") <> "" Then
		sql = sql & "aeditDate" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_aeditDate"),",")
	END IF
	IF xUpForm("htx_ixCuAttach") <> "" Then
		sql = sql & "ixCuAttach" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ixCuAttach"),",")
	END IF
	IF xUpForm("htx_adesc") <> "" Then
		sql = sql & "adesc" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_adesc"),",")
	END IF
	IF xUpForm("htx_blist") <> "" Then
		sql = sql & "blist" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_blist"),",")
	END IF
	IF xUpForm("htx_listSeq") <> "" Then
		sql = sql & "listSeq" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_listSeq"),",")
	END IF
	IF checkGIPconfig("AttachmentType") and xUpForm("htx_Attachtype") <> "" Then
		sql = sql & "Attachtype" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_Attachtype"),",")
	END IF
	IF checkGIPconfig("Attachmentkind") and xUpForm("htx_AttachkindA") <> "" Then
		sql = sql & "AttachkindA" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_AttachkindA"),",")
	END IF

	
if checkGIPconfig("AttachLarge") and xUpForm("htFile_NFileType")="2" then	'----j
	SQLP = "Select mvalue from CodeMain where codeMetaId='AttachmentLarge' and mcode='1'"
	Set RSP = conn.execute(SQLP)
	sourceFilePath = RSP(0) & "\" & xUpForm("htFile_nfileName_Large")
	Set fso = CreateObject("Scripting.FileSystemObject")
	EncodeOrNot="N"
	if xUpForm("bfx_EncodeOrNot")<>"" then EncodeOrNot=xUpForm("bfx_EncodeOrNot")
	if EncodeOrNot = "N" then	'----犰W,邠dO_,Yreject
	    filespec = HTUploadPath & xUpForm("htFile_nfileName_Large")
	    if fso.FileExists(server.mappath(filespec)) then	    
%>
		<script language=vbs>
			alert "附件檔名重複!"
			window.history.back
		</script>"  
<%		response.end
	    else
	    	ofname = xUpForm("htFile_nfileName_Large")
	    	nfname = xUpForm("htFile_nfileName_Large")
 	    end if
	else				'----犰W
		ofname = xUpForm("htFile_nfileName_Large")
		fnExt = ""
		if instrRev(ofname, ".")>0 then	fnext=mid(ofname, instrRev(ofname, "."))
		tstr = now()
		nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext		 	
	end if
	'----仵袙s
	Set MyFile = fso.GetFile(sourceFilePath)
	MyFile.Copy apath + nfname
	'----FTPBz
	  if xFTPIP<>"" and xFTPID<>"" and xFTPPWD<>"" then
		fileAction="MoveFile"
		fileTarget=nfname
		fileSource=apath + nfname
		ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"",fileTarget,fileSource 	
   	  end if  	  	
	'----SQLr
	sql = sql & "NFileType,EncodeOrNot,nfileName,"
	sqlValue = sqlValue & "'"&xUpForm("htFile_NFileType")&"','"&EncodeOrNot&"'," & pkStr(nfname,",")
	'----sWImageFile
	xsql = "INSERT INTO ImageFile(newFileName, oldFileName) VALUES(" _
		& pkStr(nfname,",") & pkStr(ofname,")")
	conn.execute xsql
else					'----@諈?
  
   For Each Form In xup.Form
    If Form.IsFile Then
    if left(Form.Name,6) = "htImg_" then
	  ofname = Form.FileName
	  fnExt = ""
	  if instrRev(ofname, ".")>0 then	fnext=mid(ofname, instrRev(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(Form.Name,7) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xup.Form(Form.Name).SaveAs apath & nfname, True
	  if xFTPIP<>"" and xFTPID<>"" and xFTPPWD<>"" then
		fileAction="MoveFile"
		fileTarget=nfname
		fileSource=apath + nfname
		ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"",fileTarget,fileSource 	
   	  end if  	    	  
  	elseif left(Form.Name,7) = "htFile_" then
	  ofname = Form.FileName
	  fnExt = ""
	  if instrRev(ofname, ".")>0 then	fnext=mid(ofname, instrRev(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(Form.Name,8) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xup.Form(Form.Name).SaveAs apath & nfname, True
	  if xFTPIP<>"" and xFTPID<>"" and xFTPPWD<>"" then
		fileAction="MoveFile"
		fileTarget=nfname
		fileSource=apath + nfname
		ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"",fileTarget,fileSource 	
   	  end if  	  	  	  
  	  xsql = "INSERT INTO ImageFile(newFileName, oldFileName) VALUES(" _
  	  	& pkStr(nfname,",") & pkStr(ofname,")")
  	  conn.execute xsql
	end if
      end if	
  Next
end if
	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	
	conn.Execute(SQL)  
end sub '---- doUpdateDB() ----

sub showDoneBox()
	doneURI= ""
	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp?" & request.serverVariables("QUERY_STRING")
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
	doneStr="新增完成！"
<%if FTPErrorMSG<>"" then%>
	doneStr="新增完成！"+VBCRLF+VBCRLF+"<%=FTPErrorMSG%>"

<%end if%>	
	alert(doneStr)
	document.location.href = "<%=doneURI%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
