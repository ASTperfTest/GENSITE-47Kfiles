<% Response.Expires = 0
HTProgCap="課程"
HTProgFunc="編修"
HTProgCode="PA001"
HTProgPrefix="paAct" %>
<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
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
	SQL = "DELETE FROM ppAct WHERE actID=" & pkStr(request.queryString("actID"),"")
	conn.Execute SQL
	showDoneBox("資料刪除成功！")

else
	EditInBothCase()
end if


sub EditInBothCase
	sqlCom = "SELECT * FROM ppAct WHERE actID=" & pkStr(request.queryString("actID"),"")
	Set RSreg = Conn.execute(sqlcom)
	pKey = ""
	pKey = pKey & "&actID=" & RSreg("actID")
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
	reg.htx_actID.value= "<%=qqRS("actID")%>"
	reg.htx_actCat.value= "<%=qqRS("actCat")%>"
	reg.htx_actName.value= "<%=qqRS("actName")%>"
	reg.htx_actDesc.value= "<%=qqRS("actDesc")%>"
	reg.htx_actTarget.value= "<%=qqRS("actTarget")%>"
	reg.htx_detailDesc.value= "<%=qqRS("detailDesc")%>"
	reg.htx_place.value= "<%=qqRS("place")%>"
	reg.htx_actTime.value= "<%=qqRS("actTime")%>"
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
  
	IF reg.htx_actID.value = Empty Then
		MsgBox replace(nMsg,"{0}","課程代碼"), 64, "Sorry!"
		reg.htx_actID.focus
		exit sub
	END IF
	IF reg.htx_actCat.value = Empty Then
		MsgBox replace(nMsg,"{0}","課程類別"), 64, "Sorry!"
		reg.htx_actCat.focus
		exit sub
	END IF
	IF reg.htx_actName.value = Empty Then
		MsgBox replace(nMsg,"{0}","項目名稱"), 64, "Sorry!"
		reg.htx_actName.focus
		exit sub
	END IF
	IF reg.htx_actDesc.value = Empty Then
		MsgBox replace(nMsg,"{0}","內容"), 64, "Sorry!"
		reg.htx_actDesc.focus
		exit sub
	END IF
	IF blen(reg.htx_actDesc.value) > 1000 Then
		MsgBox replace(replace(lMsg,"{0}","內容"),"{1}","1000"), 64, "Sorry!"
		reg.htx_actDesc.focus
		exit sub
	END IF
	IF reg.htx_actTarget.value = Empty Then
		MsgBox replace(nMsg,"{0}","對象"), 64, "Sorry!"
		reg.htx_actTarget.focus
		exit sub
	END IF
        IF reg.htx_actTime.value = Empty Then
		MsgBox replace(nMsg,"{0}","時間"), 64, "Sorry!"
		reg.htx_actTime.focus
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

<!--#include file="paActFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;"> 
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
    <title>編修表單</title>
    <link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
    </head>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
			<A href="?<%=pKey%>"></A>
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <%if (HTProgRight and 2)=2 then%>
	       <a href="<%=HTprogPrefix%>Query.asp">查詢</a>&nbsp;
	       <%end if%>	    
	       <%if (HTProgRight and 4)=4 then%>
	       <a href="<%=HTProgPrefix%>Add.asp">新增</a>&nbsp;
	       <%end if%>
	       <%if (HTProgRight and 4)=4 then%>
	       <a href="docAttach.asp?id=<% =request.queryString("actID") %>">講義</a>&nbsp;
	       <%end if%>
	       <a href="Javascript:window.history.back();">回前頁</a>
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
	sql = "UPDATE ppAct SET "
		sql = sql & "actCat=" & pkStr(request("htx_actCat"),",")
		sql = sql & "actName=" & pkStr(request("htx_actName"),",")
		sql = sql & "actDesc=" & pkStr(request("htx_actDesc"),",")
		sql = sql & "actTarget=" & pkStr(request("htx_actTarget"),",")
		sql = sql & "detailDesc=" & pkStr(request("htx_detailDesc"),",")
		sql = sql & "place=" & pkStr(request("htx_place"),",")
		sql = sql & "actTime=" & pkStr(request("htx_actTime"),",")
	sql = left(sql,len(sql)-1) & " WHERE actID=" & pkStr(request.queryString("actID"),"")

	conn.Execute(SQL)  
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=big5">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>編修表單</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
	    document.location.href="<%=HTprogPrefix%>List.asp?<%=Session("QueryPage_No")%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
