<%@ CodePage = 65001 %>
<% Response.Expires = 0
Server.ScriptTimeOut = 3000
HTProgCap="ƪ"
HTProgFunc="sW"
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
<title>sW</title>
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
'----ftpѼƳBz
	FTPErrorMSG=""
	FTPfilePath="public/Attachment"
	SQLP = "Select * from UpLoadSite where UpLoadSiteID='file'"
	Set RSP = conn.execute(SQLP)
	if not RSP.EOF  then
   		xFTPIP = RSP("UpLoadSiteFTPIP")
   		xFTPPort = RSP("UpLoadSiteFTPPort")
   		xFTPID = RSP("UpLoadSiteFTPID")
   		xFTPPWD = RSP("UpLoadSiteFTPPWD")
   	end if
'----ftpѼƳBzend 
 apath=server.mappath(HTUploadPath) & "\"
 'response.write apath & "<HR>"
  set xup = Server.CreateObject("UpDownExpress.FileUpload")
' response.end
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
	reg.htx_xiCuItem.value= "<%=request("iCuItem")%>"
	reg.htx_aEditor.value= "<%=session("userID")%>"
	reg.htx_aEditDate.value= "<%=date()%>"
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
	IF reg.htx_aTitle.value = Empty Then 
		MsgBox replace(nMsg,"{0}","W"), 64, "Sorry!"
		reg.htx_aTitle.focus
		exit sub
	END IF
	IF blen(reg.htx_aTitle.value) > 80 Then
		MsgBox replace(replace(lMsg,"{0}","W"),"{1}","80"), 64, "Sorry!"
		reg.htx_aTitle.focus
		exit sub
	END IF
	IF (reg.htx_xiCuItem.value <> "") AND (NOT isNumeric(reg.htx_xiCuItem.value)) Then
		MsgBox replace(iMsg,"{0}","ID"), 64, "Sorry!"
		reg.htx_xiCuItem.focus
		exit sub
	END IF
	IF (reg.htx_aEditDate.value <> "") AND (NOT isDate(reg.htx_aEditDate.value)) Then
		MsgBox replace(dMsg,"{0}","WǤ"), 64, "Sorry!"
		reg.htx_aEditDate.focus
		exit sub
	END IF
	IF (reg.htx_ixCuAttach.value <> "") AND (NOT isNumeric(reg.htx_ixCuAttach.value)) Then
		MsgBox replace(iMsg,"{0}","ID"), 64, "Sorry!"
		reg.htx_ixCuAttach.focus
		exit sub
	END IF
	IF blen(reg.htx_aDesc.value) > 800 Then
		MsgBox replace(replace(lMsg,"{0}",""),"{1}","800"), 64, "Sorry!"
		reg.htx_aDesc.focus
		exit sub
	END IF
	<%if checkGIPconfig("AttachLarge") then%>
	  	if reg.htFile_NFileType(0).checked = true then
		  	IF reg.htfile_NFileName.value = Empty Then 
				MsgBox replace(nMsg,"{0}","tɦW"), 64, "Sorry!"
				reg.htfile_NFileName.focus
				exit sub
			END IF
	  	elseif reg.htFile_NFileType(1).checked = true then
		  	IF reg.htFile_NFileName_Large.value = Empty Then 
				MsgBox replace(nMsg,"{0}","tɦW"), 64, "Sorry!"
				reg.htFile_NFileName_Large.focus
				exit sub
			END IF
	  	end if
	<%else%>
		IF reg.htfile_NFileName.value = Empty Then 
			MsgBox replace(nMsg,"{0}","tɦW"), 64, "Sorry!"
			reg.htfile_NFileName.focus
			exit sub
		END IF
	
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
'	SQL = "Select * From Client Where ClientID = '"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "uȤsv!!Эsإ߫Ȥs!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

sub doUpdateDB()
	sql = "INSERT INTO CuDTAttach("
	sqlValue = ") VALUES("
	IF xUpForm("htx_aTitle") <> "" Then
		sql = sql & "aTitle" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_aTitle"),",")
	END IF
	IF xUpForm("htx_xiCuItem") <> "" Then
		sql = sql & "xiCuItem" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xiCuItem"),",")
	END IF
	IF xUpForm("htx_aEditor") <> "" Then
		sql = sql & "aEditor" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_aEditor"),",")
	END IF
	IF xUpForm("htx_aEditDate") <> "" Then
		sql = sql & "aEditDate" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_aEditDate"),",")
	END IF
	IF xUpForm("htx_ixCuAttach") <> "" Then
		sql = sql & "ixCuAttach" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ixCuAttach"),",")
	END IF
	IF xUpForm("htx_aDesc") <> "" Then
		sql = sql & "aDesc" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_aDesc"),",")
	END IF
	IF xUpForm("htx_bList") <> "" Then
		sql = sql & "bList" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_bList"),",")
	END IF
	IF xUpForm("htx_listSeq") <> "" Then
		sql = sql & "listSeq" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_listSeq"),",")
	END IF
	IF checkGIPconfig("AttachmentType") and xUpForm("htx_Attachtype") <> "" Then
		sql = sql & "Attachtype" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_Attachtype"),",")
	END IF
	IF checkGIPconfig("AttachmentType") and xUpForm("htx_AttachkindA") <> "" Then
		sql = sql & "AttachkindA" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_AttachkindA"),",")
	END IF

	
if checkGIPconfig("AttachLarge") and xUpForm("htFile_NFileType")="2" then	'----j
	SQLP = "Select mValue from CodeMain where codeMetaID='AttachmentLarge' and mCode='1'"
	Set RSP = conn.execute(SQLP)
	sourceFilePath = RSP(0) & "\" & xUpForm("htFile_NFileName_Large")
	Set fso = CreateObject("Scripting.FileSystemObject")
	EncodeOrNot="N"
	if xUpForm("bfx_EncodeOrNot")<>"" then EncodeOrNot=xUpForm("bfx_EncodeOrNot")
	if EncodeOrNot = "N" then	'----ɦW,ˬdO_,Yreject
	    filespec = HTUploadPath & xUpForm("htFile_NFileName_Large")
	    if fso.FileExists(server.mappath(filespec)) then	    
%>
		<script language=vbs>
			alert "ɦW!"
			window.history.back
		</script>"  
<%		response.end
	    else
	    	ofname = xUpForm("htFile_NFileName_Large")
	    	nfname = xUpForm("htFile_NFileName_Large")
 	    end if
	else				'----ɦW
		ofname = xUpForm("htFile_NFileName_Large")
		fnExt = ""
		if instrRev(ofname, ".")>0 then	fnext=mid(ofname, instrRev(ofname, "."))
		tstr = now()
		nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext		 	
	end if
	'----ɮ׽ƻs
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
	sql = sql & "NFileType,EncodeOrNot,NFileName,"
	sqlValue = sqlValue & "'"&xUpForm("htFile_NFileType")&"','"&EncodeOrNot&"'," & pkStr(nfname,",")
	'----sWimageFile
	xsql = "INSERT INTO imageFile(newFileName, oldFileName) VALUES(" _
		& pkStr(nfname,",") & pkStr(ofname,")")
	conn.execute xsql
else					'----@몫
  For each xatt in xup.Attachments
    if left(xatt.Name,6) = "htImg_" then
	  ofname = xatt.FileName
	  fnExt = ""
	  if instrRev(ofname, ".")>0 then	fnext=mid(ofname, instrRev(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(xatt.Name,7) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xatt.SaveFile apath & nfname, false
	  if xFTPIP<>"" and xFTPID<>"" and xFTPPWD<>"" then
		fileAction="MoveFile"
		fileTarget=nfname
		fileSource=apath + nfname
		ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"",fileTarget,fileSource 	
   	  end if  	    	  
  	elseif left(xatt.Name,7) = "htFile_" then
	  ofname = xatt.FileName
	  fnExt = ""
	  if instrRev(ofname, ".")>0 then	fnext=mid(ofname, instrRev(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(xatt.Name,8) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xatt.SaveFile apath & nfname, false
	  if xFTPIP<>"" and xFTPID<>"" and xFTPPWD<>"" then
		fileAction="MoveFile"
		fileTarget=nfname
		fileSource=apath + nfname
		ftpDo xFTPIP,xFTPPort,xFTPID,xFTPPWD,fileAction,FTPfilePath,"",fileTarget,fileSource 	
   	  end if  	  	  	  
  	  xsql = "INSERT INTO imageFile(newFileName, oldFileName) VALUES(" _
  	  	& pkStr(nfname,",") & pkStr(ofname,")")
  	  conn.execute xsql
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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>sW{</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	doneStr="sWI"
<%if FTPErrorMSG<>"" then%>
	doneStr="sWI"+VBCRLF+VBCRLF+"<%=FTPErrorMSG%>"

<%end if%>	
	alert(doneStr)
	document.location.href = "<%=doneURI%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
