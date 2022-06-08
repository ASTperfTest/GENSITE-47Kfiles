<%@ CodePage = 65001 %>
<%
Response.Expires = 0
response.charset="utf-8"
HTProgCap="主題單元管理"
HTProgFunc="編修"
HTUploadPath="/public/"
HTProgCode="GE1T11"
HTProgPrefix="CtUnit" %>
<!--#include virtual = "/inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
<title>s</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
Dim xCount
taskLable="s" & HTProgCap

 apath=server.mappath(HTUploadPath) & "\"
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

if xUpForm("submitTask") = "UPDATE" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("資料更新成功！")
	end if

elseif xUpForm("submitTask") = "DELETE" then
	SQL = "DELETE FROM CtUnit WHERE ctUnitId=" & pkStr(request.queryString("ctUnitId"),"")
	conn.Execute SQL
	showDoneBox("資料刪除成功！")

else
	EditInBothCase()
end if


sub EditInBothCase
	SQLCount = "Select count(*) from CuDtGeneric where ictunit=" & pkStr(request.queryString("ctUnitId"),"")
	Set RSCount=conn.execute(SQLCount)
	xCount=RSCount(0)
	sqlcom = "SELECT htx.*,BD.sbaseDsdname,BD.rdsdcat " & _
		"FROM CtUnit AS htx Left Join BaseDsd BD ON htx.ibaseDsd=BD.ibaseDsd " & _
		"WHERE htx.ctUnitId=" & pkStr(request.queryString("ctUnitId"),"")
'response.write sqlcom
'response.end
		Set RSreg = Conn.execute(sqlcom)
	pKey = ""
	pKey = pKey & "&ctUnitId=" & RSreg("ctUnitId")
	if pKey<>"" then  pKey = mid(pKey,2)

	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
end sub

'response.write "xx="&xCount
function qqRS(fldName)
	if request("submitTask")="" then
		xValue = RSreg(fldname)
	else
		xValue = ""
		if request("htx_"&fldName) <> "" then
			xValue = request("htx_"&fldName)
		end if
	end if
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end if
end function %>

<%Sub initForm() %>
      <script language=vbs>
      sub resetForm 
	       reg.reset()
	       clientInitForm
      end sub

     sub clientInitForm
	<%if checkGIPconfig("MMOFolder") and RSreg("rdsdcat")="MMO" then %>
		document.all("MMOFolder").style.display="none"
<%	end if %>		
	reg.htx_ctUnitName.value= "<%=qqRS("ctUnitName")%>"
	reg.htx_ctUnitId.value= "<%=qqRS("ctUnitId")%>"
	initImgFile "CtUnitLogo","<%=qqRS("CtUnitLogo")%>"
	reg.htx_ctUnitKind.value= "<%=qqRS("ctUnitKind")%>"
	reg.htx_redirectUrl.value= "<%=qqRS("redirectUrl")%>"
	reg.htx_newWindow.value= "<%=qqRS("newWindow")%>"
	reg.htx_ibaseDsd.value= "<%=qqRS("ibaseDsd")%>"
	reg.xrefDSDCat.value= "<%=qqRS("rdsdcat")%>"
	reg.iBaseDSDSelect.value =  "<%=qqRS("ibaseDsd")%>---<%=qqRS("rdsdcat")%>"
	reg.htx_fctUnitOnly.value= "<%=qqRS("fctUnitOnly")%>"
	reg.htx_checkYn.value= "<%=qqRS("checkYn")%>"
	reg.htx_deptId.value= "<%=qqRS("deptId")%>"
	reg.htx_headerPart.value= "<%=qqRS("headerPart")%>"
	reg.htx_footerPart.value= "<%=qqRS("footerPart")%>"
<% 	if checkGIPconfig("CtUnitExpireCheck") then %>
	reg.htx_CtUnitExpireDay.value= "<%=qqRS("CtUnitExpireDay")%>"
<%	end if %>
<% 	if checkGIPconfig("MMOFolder") and RSreg("rdsdcat")="MMO" then %>
	    if reg.xrefDSDCat.value="MMO" then 
	        document.all("MMOFolder").style.display=""
		reg.htx_MMOFolderID.value= "<%=qqRS("mmofolderID")%>"
	    end if
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

    sub initOtherRadio(xname,value, otherName)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    	if value="" then	exit sub
		reg.all(xname).item(reg.all(xname).length-1).checked = true
		reg.all(otherName).value = value
		reg.all(xname).item(reg.all(xname).length-1).value = value
    end sub

    sub initCheckbox(xname,ckValue)
    	value = ckValue & ","
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    		end if
    	next
    end sub
    
    sub initOtherCheckbox(xname,ckValue,otherName)
    	valueArray = split(ckValue,", ")
    	valueCount = ubound(valueArray) + 1
    	value = ckValue & ","
    	ckCount = 0
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    			ckCount = ckCount + 1
    		end if
    	next
		if ckCount <> valueCount then
			reg.all(xname).item(reg.all(xname).length-1).checked = true
			reg.all(otherName).value = valueArray(ubound(valueArray))
			reg.all(xname).item(reg.all(xname).length-1).value = valueArray(ubound(valueArray))
		end if
    end sub

    sub initImgFile(xname, value)
		reg.all("htImgActCK_"&xname).value=""
		reg.all("htImg_"&xname).style.display="none"
		reg.all("hoImg_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).src= "<%=HTUploadPath%>" & value
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub

Sub addLogo(xname)	'sWlogo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="addLogo"
End sub

Sub chgLogo(xname)	'logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="editLogo"
End sub

Sub orgLogo(xname)	'
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value=""
End sub

Sub delLogo(xname)	'Rlogo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htImg_"&xname).value=""
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value="delLogo"
End sub

    sub initAttFile(xname, value, orgValue)
		reg.all("htFileActCK_"&xname).value=""
		reg.all("htFile_"&xname).style.display="none"
		reg.all("hoFile_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).innerText= orgValue
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub

Sub addXFile(xname)	'sWlogo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="addLogo"
End sub

Sub chgXFile(xname)	'logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="editLogo"
End sub

Sub orgXFile(xname)	'
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value=""
End sub

Sub delXFile(xname)	'Rlogo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htFile_"&xname).value=""
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value="delLogo"
End sub

 </script>	
<%End sub '---- initForm() ----%>


<%Sub showForm()  	'===================== Client Side Validation Put HERE =========== %>
      <script language="vbscript">
      sub formModSubmit()
    
'----- SubmitedXbo ApU not valid  exit sub }------------
'msg1 = "gusvAoI"
'msg2 = "guWvAoI"
'
'  If reg.tfx_ClientName.value = Empty Then
'     MsgBox msg2, 64, "Sorry!"
'     settab 0
'     reg.tfx_ClientName.focus
'     Exit Sub
'  End if
	nMsg = "gu{0}vAoI"
	lMsg = "u{0}vh{1}I"
	dMsg = "u{0}v yyyy/mm/dd I"
	iMsg = "u{0}vI"
	pMsg = "u{0}v " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" @I"
  
	IF blen(reg.htx_ctUnitName.value) > 200 Then
		MsgBox replace(replace(lMsg,"{0}","DDW"),"{1}","200"), 64, "Sorry!"
		reg.htx_ctUnitName.focus
		exit sub
	END IF
	IF (reg.htx_ctUnitId.value <> "") AND (NOT isNumeric(reg.htx_ctUnitId.value)) Then
		MsgBox replace(iMsg,"{0}","DDID"), 64, "Sorry!"
		reg.htx_ctUnitId.focus
		exit sub
	END IF
	IF reg.htImg_CtUnitLogo.value <> "" Then
		xIMGname = reg.htImg_CtUnitLogo.value
		xFileType = ""
		if instr(xIMGname, ".")>0 then	xFileType=lcase(mid(xIMGname, instr(xIMGname, ".")+1))
		IF xFileType<>"gif" AND xFileType<>"jpg" AND xFileType<>"jpeg" then
			MsgBox replace(pMsg,"{0}","LOGO"), 64, "Sorry!"
			reg.htImg_CtUnitLogo.focus
			exit sub
		END IF
	END IF
	IF reg.htx_ctUnitKind.value = Empty Then 
		MsgBox replace(nMsg,"{0}",""), 64, "Sorry!"
		reg.htx_ctUnitKind.focus
		exit sub
	END IF
	IF blen(reg.htx_redirectUrl.value) > 100 Then
		MsgBox replace(replace(lMsg,"{0}","URL"),"{1}","100"), 64, "Sorry!"
		reg.htx_redirectUrl.focus
		exit sub
	END IF
	IF (reg.htx_ibaseDsd.value <> "") AND (NOT isNumeric(reg.htx_ibaseDsd.value)) Then
		MsgBox replace(iMsg,"{0}","wq"), 64, "Sorry!"
		reg.htx_ibaseDsd.focus
		exit sub
	END IF
	IF reg.htx_fctUnitOnly.value = Empty Then 
		MsgBox replace(nMsg,"{0}","u"), 64, "Sorry!"
		reg.htx_fctUnitOnly.focus
		exit sub
	END IF
<% 	if checkGIPconfig("CtUnitExpireCheck") then %>
	IF (reg.htx_CtUnitExpireDay.value <> "") AND (NOT isNumeric(reg.htx_CtUnitExpireDay.value)) Then
		MsgBox replace(iMsg,"{0}","O"), 64, "Sorry!"
		reg.htx_CtUnitExpireDay.focus
		exit sub
	END IF
<%	end if %>
	IF (reg.htx_headerPart.value <> "") AND (NOT isNumeric(reg.htx_headerPart.value)) Then
		MsgBox replace(iMsg,"{0}","YID"), 64, "Sorry!"
		reg.htx_headerPart.focus
		exit sub
	END IF
	IF (reg.htx_footerPart.value <> "") AND (NOT isNumeric(reg.htx_footerPart.value)) Then
		MsgBox replace(iMsg,"{0}","ID"), 64, "Sorry!"
		reg.htx_footerPart.focus
		exit sub
	END IF  
	<%if checkGIPconfig("MMOFolder") and RSreg("rdsdcat")="MMO" then %>
	    if reg.xrefDSDCat.value="MMO" and reg.htx_MMOFolderID.value="0" then
		alert "IhCDDs!"
		reg.htx_MMOFolderID.focus
		exit sub		
	    end if	
	<%end if%>    
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub

   sub formDelSubmit()
         'chky=msgbox("`NI"& vbcrlf & vbcrlf &"@ATwRH"& vbcrlf , 48+1, "`NII")
         chky=msgbox("注意！"& vbcrlf & vbcrlf &"　你確定刪除資料嗎？"& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
	       reg.submitTask.value = "DELETE"
	      reg.Submit
       end If
  end sub
  
sub iBaseDSDSpilt()
	if reg.iBaseDSDSelect.value="" then 
		reg.htx_iBaseDSD.value=""
		reg.xrefDSDCat.value=""
		<%if checkGIPconfig("MMOFolder") and RSreg("rdsdcat")="MMO" then %>
			document.all("MMOFolder").style.display="none"
		        reg.htx_MMOFolderID.value="0"
		<%end if%>
		exit sub
	end if
	BaseDSDPos=instr(reg.iBaseDSDSelect.value,"---")
	reg.htx_iBaseDSD.value=Left(reg.iBaseDSDSelect.value,BaseDSDPos-1)
	reg.xrefDSDCat.value=mid(reg.iBaseDSDSelect.value,BaseDSDPos+3)
	<%if checkGIPconfig("MMOFolder") and RSreg("rdsdcat")="MMO" then %>
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
		alert "s!"
		reg.htx_MMOFolderID.focus
		exit sub				
	else
		window.open "/MMO/MMOAddFolder.asp?MMOFolderID="&reg.htx_MMOFolderID.value,"","height=400,width=550"
	end if
end sub
sub MMOFolderAddReload()
	document.location.reload()
end sub
<%end if%>
</script>

<!--#include file="CtUnitFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
		<%if (HTProgRight and 1)=1 then%>
			<A href="CtUnitQuery.asp" title="重設查詢條件">查詢</A>
		<%end if%>
		<%if (HTProgRight and 8)=8 then%>
			<A href="CTUnitDSD.asp?ctUnit=<%=RSreg("ctUnitId")%>&baseDSD=<%=RSreg("ibaseDsd")%>">DSD</A>
		<%end if%>		
		<%if (HTProgRight and 8)=8 then%>
			<A href="CTDSDSwap.asp?ctUnit=<%=RSreg("ctUnitId")%>&baseDSD=<%=RSreg("ibaseDsd")%>&phase=swap">變換範本</A>
		<%end if%>		
		<%if (HTProgRight and 4)=4 then%>
			<A href="CtUnitAdd.asp?phase=add" title="新增">新增</A>
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
	    <td align=center colspan=2 width=80% height=230 valign=top>            
<%End sub '--- showHTMLHead() ------%>


<%Sub ShowHTMLTail() %>     
    </td>
  </tr>  
</table> 
</body>
</html>
<%End sub '--- showHTMLTail() ------%>


<%Sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
'  			window.history.back
	</script>
<%
    End sub '---- showErrBox() ----

Sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- d{Xbo	ApUA] errMsg="xxx"  exit sub ------
'	if Request("tfx_TaxID") <> "" Then
'	  SQL = "Select * From Client Where TaxID = '"& Request("tfx_TaxID") &"'"
'	  set RSreg = conn.Execute(SQL)
'	  if not RSreg.EOF then
'		if trim(RSreg("ClientId")) <> request("pfx_ClientID") Then
'			errMsg = "u@sv!!sJ@s!"
'			exit sub
'		end if
'	  end if
'	end if

End sub '---- checkDBValid() ----

Sub doUpdateDB()
	sql = "UPDATE CtUnit SET "
		sql = sql & "ctUnitName=" & pkStr(xUpForm("htx_ctUnitName"),",")
	IF xUpForm("htImgActCK_CtUnitLogo") <> "" Then
	  actCK = xUpForm("htImgActCK_CtUnitLogo")
	  if actCK="editLogo" OR actCK="addLogo" then
		fname = ""
		For Each Form In xup.Form
		If Form.IsFile Then 
		  if Form.Name = "htImg_CtUnitLogo" then
			ofname = Form.FileName
			fnExt = ""
			if instr(ofname, ".")>0 then	fnext=mid(ofname, instr(ofname, "."))
			tstr = now()
			nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
			sql = sql & "CtUnitLogo=" & pkStr(nfname,",")
			Set fso = server.CreateObject("Scripting.FileSystemObject")
			IF xUpForm("hoImg_CtUnitLogo") <> "" Then _
				fso.DeleteFile apath & xUpForm("hoImg_CtUnitLogo")
			xup.Form(Form.Name).SaveAs apath & nfname, True
		  end if
		 end if 
		Next
	  elseif actCK="delLogo" then
		xup.DeleteFile apath & xUpForm("hoImg_CtUnitLogo")
		sql = sql & "CtUnitLogo=null,"
	  end if
	END IF

 	if checkGIPconfig("CtUnitExpireCheck") then
		sql = sql & "CtUnitExpireDay=" & pkStr(xUpForm("htx_CtUnitExpireDay"),",")
	end if
 	if checkGIPconfig("MMOFolder") then
 	    	if xUpForm("htx_MMOFolderID")<>"0" then
			sql = sql & "mmofolderId=" & pkStr(xUpForm("htx_MMOFolderID"),",")
		else
			sql = sql & "mmofolderId=null,"
		end if
	end if
		sql = sql & "ctUnitKind=" & pkStr(xUpForm("htx_ctUnitKind"),",")
		sql = sql & "redirectUrl=" & pkStr(xUpForm("htx_redirectUrl"),",")
		sql = sql & "newWindow=" & pkStr(xUpForm("htx_newWindow"),",")
		sql = sql & "ibaseDsd=" & pkStr(xUpForm("htx_ibaseDsd"),",")
		sql = sql & "headerPart=" & pkStr(xUpForm("htx_headerPart"),",")
		sql = sql & "footerPart=" & pkStr(xUpForm("htx_footerPart"),",")
		sql = sql & "fctUnitOnly=" & pkStr(xUpForm("htx_fctUnitOnly"),",")
		sql = sql & "checkYn=" & pkStr(xUpForm("htx_checkYn"),",")
		sql = sql & "deptId=" & pkStr(xUpForm("htx_deptId"),",")
	sql = left(sql,len(sql)-1) & " WHERE ctUnitId=" & pkStr(request.queryString("ctUnitId"),"")

	conn.Execute(SQL)  
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>s</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
	    document.location.href="<%=HTprogPrefix%>List.asp?<%=Session("QueryPage_No")%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
