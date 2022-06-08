<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap="單元資料定義"
HTProgFunc="編修欄位"
HTUploadPath="/public/"
HTProgCode="GE1T01"
HTProgPrefix="bDSDField" %>
<!--#include virtual = "/inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
<title>編修表單</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
taskLable="編輯" & HTProgCap

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
	SQL = "DELETE FROM BaseDsdfield WHERE ibaseField=" & pkStr(request.queryString("iBaseField"),"")
	conn.Execute SQL
	showDoneBox("資料刪除成功！")

else
	EditInBothCase()
end if


sub EditInBothCase
	sqlCom = "SELECT htx.* FROM BaseDsdfield AS htx WHERE htx.ibaseField=" & pkStr(request.queryString("iBaseField"),"")
	Set RSreg = Conn.execute(sqlcom)
	pKey = ""
	pKey = pKey & "&ibaseField=" & RSreg("iBaseField")
	if pKey<>"" then  pKey = mid(pKey,2)

	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
end sub


function qqRS(fldName)
	if xUpForm("submitTask")="" then
		xValue = RSreg(fldname)
	else
		xValue = ""
		if xUpForm("htx_"&fldName) <> "" then
			xValue = xUpForm("htx_"&fldName)
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
	reg.htx_xfieldSeq.value= "<%=qqRS("xfieldSeq")%>"
	reg.htx_iBaseDSD.value= "<%=qqRS("iBaseDSD")%>"
	reg.htx_xfieldName.value= "<%=qqRS("xfieldName")%>"
	reg.htx_iBaseField.value= "<%=qqRS("ibaseField")%>"
	reg.htx_inUse.value= "<%=qqRS("inUse")%>"
	reg.htx_xfieldLabel.value= "<%=qqRS("xfieldLabel")%>"
	reg.htx_xfieldDesc.value= "<%=qqRS("xfieldDesc")%>"
	reg.htx_xdataType.value= "<%=qqRS("xdataType")%>"
	reg.htx_xdataLen.value= "<%=qqRS("xdataLen")%>"
	initRadio "htx_xcanNull","<%=qqRS("xcanNull")%>"
	reg.htx_xinputType.value= "<%=qqRS("xinputType")%>"
	reg.htx_xrefLookup.value= "<%=qqRS("xrefLookup")%>"
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

Sub addLogo(xname)	'新增logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="addLogo"
End sub

Sub chgLogo(xname)	'更換logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="editLogo"
End sub

Sub orgLogo(xname)	'原圖
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value=""
End sub

Sub delLogo(xname)	'刪除logo
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

Sub addXFile(xname)	'新增logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="addLogo"
End sub

Sub chgXFile(xname)	'更換logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="editLogo"
End sub

Sub orgXFile(xname)	'原圖
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value=""
End sub

Sub delXFile(xname)	'刪除logo
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
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
'msg1 = "請務必填寫「客戶編號」，不得為空白！"
'msg2 = "請務必填寫「客戶名稱」，不得為空白！"
'
'  If reg.tfx_ClientName.value = Empty Then
'     MsgBox msg2, 64, "Sorry!"
'     settab 0
'     reg.tfx_ClientName.focus
'     Exit Sub
'  End if
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
  
	IF (reg.htx_iBaseDSD.value <> "") AND (NOT isNumeric(reg.htx_iBaseDSD.value)) Then
		MsgBox replace(iMsg,"{0}","單元資料定義ID"), 64, "Sorry!"
		reg.htx_iBaseDSD.focus
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
  
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub

   sub formDelSubmit()
         chky=msgbox("注意！"& vbcrlf & vbcrlf &"　你確定刪除資料嗎？"& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
	       reg.submitTask.value = "DELETE"
	      reg.Submit
       end If
  end sub

</script>

<!--#include file="bDSDFieldFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
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

'---- 後端資料驗證檢查程式碼放在這裡	，如下例，有錯時設 errMsg="xxx" 並 exit sub ------
'	if Request("tfx_TaxID") <> "" Then
'	  SQL = "Select * From Client Where TaxID = '"& Request("tfx_TaxID") &"'"
'	  set RSreg = conn.Execute(SQL)
'	  if not RSreg.EOF then
'		if trim(RSreg("ClientId")) <> request("pfx_ClientID") Then
'			errMsg = "「統一編號」重複!!請重新輸入客戶統一編號!"
'			exit sub
'		end if
'	  end if
'	end if

End sub '---- checkDBValid() ----

Sub doUpdateDB()
	sql = "UPDATE BaseDsdfield SET "
		sql = sql & "xfieldSeq=" & pkStr(xUpForm("htx_xfieldSeq"),",")
		sql = sql & "xfieldName=" & pkStr(xUpForm("htx_xfieldName"),",")
		sql = sql & "inUse=" & pkStr(xUpForm("htx_inUse"),",")
		sql = sql & "xfieldLabel=" & pkStr(xUpForm("htx_xfieldLabel"),",")
		sql = sql & "xfieldDesc=" & pkStr(xUpForm("htx_xfieldDesc"),",")
		sql = sql & "xdataType=" & pkStr(xUpForm("htx_xdataType"),",")
		sql = sql & "xdataLen=" & pkStr(xUpForm("htx_xdataLen"),",")
		sql = sql & "xcanNull=" & pkStr(xUpForm("htx_xcanNull"),",")
		sql = sql & "xinputType=" & pkStr(xUpForm("htx_xinputType"),",")
		sql = sql & "xrefLookup=" & pkStr(xUpForm("htx_xrefLookup"),",")
	sql = left(sql,len(sql)-1) & " WHERE ibaseField=" & pkStr(request.queryString("iBaseField"),"")

	conn.Execute(SQL)  
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) 
	mpKey = ""
	mpKey = mpKey & "&iBaseDSD=" & xUpForm("htx_iBaseDSD")
	if mpKey<>"" then  mpKey = mid(mpKey,2)
	doneURI= "BaseDSDEditList.asp"
	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp"
%>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>編修表單</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
	    document.location.href="<%=doneURI%>?<%=mpKey%>&phase=edit"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
