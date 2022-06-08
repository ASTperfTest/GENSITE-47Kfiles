<%@ CodePage = 65001 %>
<%HTProgCap="代碼定義"
HTProgCode="Pn50M03"
HTProgPrefix="CodeTable" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
Dim RSreg
Dim formFunction
taskLable="查詢" & HTProgCap

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
	SQL = "DELETE FROM CodeMetaDef WHERE codeId=" & pkStr(request.queryString("codeId"),"")
'response.write SQL & "<BR>"
	conn.Execute SQL
	if request("tfx_codeTblName")="CodeMain" then
		SQL = "Delete from " & request("tfx_codeTblName") & " Where codeMetaId=N'" & request("pfx_codeId") & "'"
		conn.execute(SQL)
	end if
	showDoneBox("資料刪除成功！")

else
	EditInBothCase()
end if

sub EditInBothCase
	showHTMLHead()
	sqlCom = "SELECT * FROM CodeMetaDef WHERE codeId=" & pkStr(request.queryString("codeId"),"")
  Set RSreg = Conn.execute(sqlcom)

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
		if request("pfx_"&fldName) <> "" then
			xValue = request("pfx_"&fldName)
		elseif request("tfx_"&fldName) <> "" then
			xValue = request("tfx_"&fldName)
		elseif request("dfx_"&fldName) <> "" then
			xValue = request("dfx_"&fldName)
		elseif request("sfx_"&fldName) <> "" then
			xValue = request("sfx_"&fldName)
		elseif request("nfx_"&fldName) <> "" then
			xValue = request("nfx_"&fldName)
		elseif request("bfx_"&fldName) <> "" then
			xValue = request("bfx_"&fldName)
		end if
	end if
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end if
end function %>

<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm
	reg.pfx_codeId.value = "<%=qqRS("codeId")%>"
	reg.tfx_codeName.value = "<%=qqRS("codeName")%>"
	reg.tfx_codeTblName.value = "<%=qqRS("codeTblName")%>"
	reg.tfx_CodeSrcFld.value = "<%=qqRS("CodeSrcFld")%>"
	reg.tfx_CodeSrcItem.value = "<%=qqRS("CodeSrcItem")%>"
	reg.tfx_CodeValueFld.value = "<%=qqRS("CodeValueFld")%>"
	reg.tfx_CodeDisplayFld.value = "<%=qqRS("CodeDisplayFld")%>"
	reg.tfx_CodeSortFld.value = "<%=qqRS("CodeSortFld")%>"
	reg.tfx_CodeValueFldName.value = "<%=qqRS("CodeValueFldName")%>"
	reg.tfx_CodeDisplayFldName.value = "<%=qqRS("CodeDisplayFldName")%>"
	reg.tfx_CodeSortFldName.value = "<%=qqRS("CodeSortFldName")%>"
	reg.tfx_CodeType.value = "<%=qqRS("CodeType")%>"
	reg.tfx_CodeRank.value = "<%=qqRS("CodeRank")%>"
	<%if RSreg("showOrNot")="Y" then%>
		reg.bfx_showOrNot.checked=true
	<%end if%>
<%  if checkGIPconfig("codeML") then 
		if RSreg("isML")="Y" then %>
			reg.bfx_isML.checked=true
<%		end if
	end if %>
	reg.tfx_CodeXMLSpec.value = "<%=qqRS("CodeXMLSpec")%>"
end sub

sub window_onLoad
	clientInitForm
end sub
</script>	
<% end sub '---- initForm() ----%>

<% sub showForm()  	'===================== Client Side Validation Put HERE =========== %>
<script language="vbscript">

Sub formModSubmit()
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
msg1 = "請務必輸入「代碼ID」，不得為空白！"
msg2 = "請務必輸入「代碼名稱」，不得為空白！"
msg3 = "請務必輸入「存放資料表」，不得為空白！"
msg4 = "請務必輸入「Value值欄」，不得為空白！"
msg5 = "請務必輸入「Display欄」，不得為空白！"
msg6 = "請務必輸入「過濾(排序)欄」，不得為空白！"

  If reg.pfx_codeId.value = Empty Then
     MsgBox msg1, 64, "Sorry!"
     reg.pfx_codeId.focus
     Exit Sub
  End if
  If reg.tfx_codeName.value = Empty Then
     MsgBox msg2, 64, "Sorry!"
     reg.tfx_codeName.focus
     Exit Sub
  End if
  If reg.tfx_codeTblName.value = Empty Then
     MsgBox msg3, 64, "Sorry!"
     reg.tfx_codeTblName.focus
     Exit Sub
  End if
  If reg.tfx_CodeValueFld.value = Empty Then
     MsgBox msg4, 64, "Sorry!"
     reg.tfx_CodeValueFld.focus
     Exit Sub
  End if
  If reg.tfx_CodeDisplayFld.value = Empty Then
     MsgBox msg5, 64, "Sorry!"
     reg.tfx_CodeDisplayFld.focus
     Exit Sub
  End if     
  reg.submitTask.value = "UPDATE"
  reg.Submit
End Sub

Sub formDelSubmit()
 xStr=""
 if ucase(reg.tfx_codeTblName.value)="CODEHT" then
 	xStr="刪除定義將一併刪除其下資料,"
 end if
 chky=msgbox("注意！"& vbcrlf & vbcrlf & xStr & "　你確定刪除此定義嗎？"& vbcrlf , 48+1, "請注意！！")
 if chky=vbok then
	reg.submitTask.value = "DELETE"
	reg.Submit
 End If
End Sub

</script>

<!--#include file="CodeTableForm.inc"-->

<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
<title>編修表單</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2"><%=Title%> </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" width="60%">【編修代碼定義】標有<font color=red>※</font>符號欄位為必填欄位</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><% if (HTProgRight and 2)=2 then %><a href="<%=HTprogPrefix%>Query.asp">查詢</a><% End IF %>&nbsp;
      <% if (HTProgRight and 4)=4 then %><a href="<%=HTprogPrefix%>Add.asp">新增</a><% End IF %>
      &nbsp;</td>
  </tr>  
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>  
  <tr>
    <td width="80%" height=230 valign=top colspan="2">
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
'  			window.history.back
	</script>
<%
end sub '---- showErrBox() ----

sub checkDBValid()	'===================== Server Side Validation Put HERE =================


end sub '---- checkDBValid() ----

sub doUpdateDB()
	xshowOrNot="N"
	if request("bfx_showOrNot")<>"" then xshowOrNot="Y"
	sql = "UPDATE CodeMetaDef SET showOrNot=N'"&xshowOrNot&"', "
	if checkGIPconfig("codeML") then
		xisML="N"
		if request("bfx_isML")<>"" then xisML="Y"
		sql = sql & "isML=N'"&xisML&"', "
	end if
	sqlWhere = ""
	for each x in request.form
	  if mid(x,2,3) = "fx_" AND left(x,1)<>"b" then
		xfldName = mid(x,5)
		select case left(x,1)
		  case "p"
			if sqlWhere="" then 
				sqlWhere = " WHERE " & mid(x,5) & "=" & pkStr(request(x),"")
			else
				sqlWhere = sqlWhere & " AND " & mid(x,5) & "=" & pkStr(request(x),"")
			end if
		  case "d"
			sql = sql & " " & mid(x,5) & "=" & pkStr(request(x),",")
		  case "n"
			sql = sql & " " & mid(x,5) & "=" & drn(x)
		  case else
			sql = sql & " " & mid(x,5) & "=" & pkStr(request(x),",")
		end select
	  end if
	next 
	sql = left(sql,len(sql)-1) & sqlWHERE
'	response.write sql
'	response.end
	conn.Execute(SQL)  
end sub '---- doUpdateDB() ----%>

<% sub showDoneBox(lMsg) %>
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
	document.location.href="<%=HTprogPrefix%>List.asp?page_no=<%=Session("QueryPage_No")%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
