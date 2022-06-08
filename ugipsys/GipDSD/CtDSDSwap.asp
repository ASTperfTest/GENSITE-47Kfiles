<%@ CodePage = 65001 %>
<%
Response.Expires = 0
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
 if request.querystring("phase")<>"swap" then
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
	sqlCom = "SELECT htx.*, b.sbaseDsdname, b.sbaseTableName " _
		& " FROM CtUnit AS htx LEFT JOIN BaseDsd AS b ON b.ibaseDsd=htx.ibaseDsd" _
		& " WHERE htx.ctUnitId=" & pkStr(request.queryString("CtUnit"),"")
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
    end sub

    sub window_onLoad
         clientInitForm
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
  
	IF (reg.new_ibaseDsd.value <> "") AND (NOT isNumeric(reg.new_ibaseDsd.value)) Then
		MsgBox replace(iMsg,"{0}","新資料範本"), 64, "Sorry!"
		reg.new_ibaseDsd.focus
		exit sub
	END IF
	IF (reg.new_ibaseDsd.value = reg.htx_ibaseDsd.value) Then
		MsgBox "範本資料相同", 64, "Sorry!"
		reg.new_ibaseDsd.focus
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

<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="CTDSDSwap.asp?ctUnit=<% =request.querystring("ctUnit") %>&baseDSD=<% =request.querystring("baseDSD") %>">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="eTableLable" align="right">主題單元名稱：</TD>
<TD class="eTableContent"> <%=RSreg("CtUnitName")%>
<input type="hidden" name="htx_ctUnitId" value="<%=RSreg("ctUnitId")%>">
<input type="hidden" name="htx_ibaseDsd" value="<%=RSreg("ibaseDsd")%>">
<input type="hidden" name="htx_sbaseTableName" value="<%=RSreg("sbaseTableName")%>">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">原資料範本：</TD>
<TD class="eTableContent"> <%=RSreg("sbaseDsdname")%>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">新資料範本：</TD>
<TD class="eTableContent"><Select name="new_ibaseDsd" size=1>
<option value="">請選擇</option>
			<%SQL="Select ibaseDsd,sbaseDsdname from BaseDsd where ibaseDsd IS NOT NULL AND inUse='Y' Order by ibaseDsd"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
</TABLE>
</CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
        <%if (HTProgRight AND 8) <> 0 then %>
               <input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()">
        <% End If %>
        
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">

 </td></tr>
</table>


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
'	if xUpForm("new_ibaseDsd") = "" Then
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

	xOrgBaseDsd = xUpForm("htx_ibaseDsd")
	ctUnitId = xUpForm("htx_ctUnitId")
	xOrgBaseTableName = xUpForm("htx_sbaseTableName")

	xNewBaseDsd = xUpForm("new_ibaseDsd")
	sql = "SELECT * FROM BaseDsd where ibaseDsd=" & xNewBaseDsd
	set RS = conn.execute(sql)
	xNewBaseTableName = RS("sbaseTableName")

	sql = "insert into " & xNewBaseTableName & "(gicuitem)" _
		& " select icuitem from CuDtGeneric" _
		& " where ictunit=" & ctUnitId
		
	sql = sql & "; " & "delete " & xOrgBaseTableName  _
		& " where giCuItem in (select icuitem from CuDtGeneric" _
		& " where ictunit=" & ctUnitId & ")"

	sql = sql & "; " & "update CuDtGeneric" _
		& " set ibaseDsd=" & xNewBaseDsd _
		& " where ictunit=" & ctUnitId

	sql = sql & "; " & "update CtUnit" _
		& " set ibaseDsd=" & xNewBaseDsd _
		& " where ctUnitId=" & ctUnitId

	'response.write sql
	'response.end
	conn.Execute(SQL)  
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) %>
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
	    document.location.href="<%=HTprogPrefix%>List.asp?<%=Session("QueryPage_No")%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
