<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="資料物件欄位管理"
HTProgFunc="編修資料物件欄位"
HTProgCode="HT011"
HTProgPrefix="HtDfield" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
taskLable="編輯" & HTProgCap

if request("submitTask") = "UPDATE" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("資料更新成功！")
	end if

elseif request("submitTask") = "DELETE" then
	SQL = "DELETE FROM HtDfield WHERE entityId=" & pkStr(request.queryString("entityId"),"") & " AND xfieldName=" & pkStr(request.queryString("xfieldName"),"")
	conn.Execute SQL
	showDoneBox("資料刪除成功！")

else
	EditInBothCase()
end if


sub EditInBothCase
	sqlCom = "SELECT * FROM HtDfield WHERE entityId=" & pkStr(request.queryString("entityId"),"") & " AND xfieldName=" & pkStr(request.queryString("xfieldName"),"")
	Set RSreg = Conn.execute(sqlcom)
	pKey = ""
	pKey = pKey & "&entityId=" & RSreg("entityId")
	pKey = pKey & "&xfieldName=" & RSreg("xfieldName")
	if pKey<>"" then  pKey = mid(pKey,2)

	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
end sub


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
	reg.htx_xfieldName.value= "<%=qqRS("xfieldName")%>"
	reg.htx_entityId.value= "<%=qqRS("entityId")%>"
	reg.htx_xfieldLabel.value= "<%=qqRS("xfieldLabel")%>"
	reg.htx_xfieldDesc.value= "<%=qqRS("xfieldDesc")%>"
	reg.htx_xfieldSeq.value= "<%=qqRS("xfieldSeq")%>"	
	reg.htx_xdataType.value= "<%=qqRS("xdataType")%>"
	reg.htx_xdataLen.value= "<%=qqRS("xdataLen")%>"
	reg.htx_xinputLen.value= "<%=qqRS("xinputLen")%>"	
	initRadio "htx_xcanNull","<%=qqRS("xcanNull")%>"
	initRadio "htx_xisPrimaryKey","<%=qqRS("xisPrimaryKey")%>"
	initRadio "htx_xidentity","<%=qqRS("xidentity")%>"
	reg.htx_xdefaultvalue.value= "<%=qqRS("xdefaultvalue")%>"
	reg.htx_xclientDefault.value= "<%=qqRS("xclientDefault")%>"
	reg.htx_xinputType.value= "<%=qqRS("xinputType")%>"
	reg.htx_xrefLookup.value= "<%=qqRS("xrefLookup")%>"	
	reg.htx_xrows.value= "<%=qqRS("xrows")%>"		
	reg.htx_xcols.value= "<%=qqRS("xcols")%>"				
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

    sub initOtherCheckbox(xname,ckValue,otherName)
    	valueArray = split(ckValue,",")
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
  
	IF reg.htx_xfieldName.value = Empty Then
		MsgBox replace(nMsg,"{0}","資料欄位名稱"), 64, "Sorry!"
		reg.htx_xfieldName.focus
		exit sub
	END IF
	IF reg.htx_entityID.value = Empty Then
		MsgBox replace(nMsg,"{0}","資料物件代碼"), 64, "Sorry!"
		reg.htx_entityID.focus
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

<!--#include file="HtDfieldFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;"> 
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
    <title>編修表單</title>
    <link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
    </head>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>i<%=HTProgFunc%>j</td>
		<td width="50%" class="FormLink" align="right">
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <a href="<%=HTProgPrefix%>List.asp?<%=pKey%>">回前頁</a>
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
	sql = "UPDATE HtDfield SET "
		sql = sql & "xfieldName=" & pkStr(request("htx_xfieldName"),",")
		sql = sql & "xfieldLabel=" & pkStr(request("htx_xfieldLabel"),",")
		sql = sql & "xfieldDesc=" & pkStr(request("htx_xfieldDesc"),",")
		sql = sql & "xfieldSeq=" & drn("htx_xfieldSeq")			
		sql = sql & "xdataType=" & pkStr(request("htx_xdataType"),",")
		sql = sql & "xdataLen=" & drn("htx_xdataLen")
		sql = sql & "xinputLen=" & drn("htx_xinputLen")		
		sql = sql & "xcanNull=" & pkStr(request("htx_xcanNull"),",")
		sql = sql & "xisPrimaryKey=" & pkStr(request("htx_xisPrimaryKey"),",")
		sql = sql & "xidentity=" & pkStr(request("htx_xidentity"),",")
		sql = sql & "xdefaultvalue=" & pkStr(request("htx_xdefaultvalue"),",")
		sql = sql & "xclientDefault=" & pkStr(request("htx_xclientDefault"),",")
		sql = sql & "xinputType=" & pkStr(request("htx_xinputType"),",")
		sql = sql & "xrefLookup=" & pkStr(request("htx_xrefLookup"),",")		
		sql = sql & "xrows=" & drn("htx_xrows")	
		sql = sql & "xcols=" & drn("htx_xcols")						
	sql = left(sql,len(sql)-1) & " WHERE entityId=" & pkStr(request.queryString("entityId"),"") & " AND xfieldName=" & pkStr(request.queryString("xfieldName"),"")

	conn.Execute(SQL)  
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) 
	mpKey = ""
	mpKey = mpKey & "&entityId=" & request("htx_entityId")
	if mpKey<>"" then  mpKey = mid(mpKey,2)
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
	    document.location.href="<%=HTprogPrefix%>List.asp?<%=mpKey%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
