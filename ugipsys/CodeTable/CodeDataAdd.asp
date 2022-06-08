<%@ CodePage = 65001 %>
<%response.expires=0
HTProgCap="代碼維護"
HTProgCode="Pn90M02"
HTProgPrefix="CodeData" %>
<!--#include virtual = "/inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Query.asp">
<title>查詢表單</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
taskLable="新增" & HTProgCap
	SQL="Select * from CodeMetaDef where codeId=N'" & session("codeID") & "'"	
	SET RSCode=conn.execute(SQL)
if request("submittask")="新增存檔" then
   if isNull(RSCode("CodeSrcFld")) then
	SQL="Select * from " & RSCode("CodeTblName") & " where " & RSCode("CodeValueFld") & "=N'" & request("tfx_Value") & "'"
   else	
	SQL="Select * from " & RSCode("CodeTblName") & " where " & RSCode("CodeSrcFld") & "=N'" & RSCode("CodeSrcItem") & "' AND " & RSCode("CodeValueFld") & "=N'" & request("tfx_Value") & "'"
   end if
   SET RSValidate=conn.execute(SQL)	   
   if not RSValidate.EOF then %>
	<script language=VBScript>
	  alert("Value值已建立，無法再次建立！")
	  window.history.back
	</script>	
<% else	
	if RSCode("CodeTblName")="PAOUNITM" then
		SQL="Insert Into " & RSCode("CodeTblName") & " (ORG," & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & "," & RSCode("CodeSortFld") & ") Values(N'310905500Q',N'" & request("tfx_value") & "',N'" &  request("tfx_display") & "',N'" & request("tfx_FldSort") & "')"			
	elseif isNull(RSCode("CodeSrcFld")) then
		if isNull(RSCode("CodeSortFld")) then
			SQL="Insert Into " & RSCode("CodeTblName") & " (" & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & ") Values(N'" & request("tfx_value") & "',N'" &  request("tfx_display") & "')"			
		else
			if request("tfx_FldSort")="" then
				SQL="Insert Into " & RSCode("CodeTblName") & " (" & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & "," & RSCode("CodeSortFld") & ") Values(N'" & request("tfx_value") & "',N'" &  request("tfx_display") & "',NULL)"			
			else
				SQL="Insert Into " & RSCode("CodeTblName") & " (" & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & "," & RSCode("CodeSortFld") & ") Values(N'" & request("tfx_value") & "',N'" &  request("tfx_display") & "',N'" & request("tfx_FldSort") & "')"		
			end if
		end if			
	else
		if isNull(RSCode("CodeSortFld")) then
			SQL="Insert Into " & RSCode("CodeTblName") & " (" & RSCode("CodeSrcFld") & "," & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & ") Values(N'" & RSCode("CodeSrcItem") & "',N'" & request("tfx_value") & "',N'" &  request("tfx_display") & "')"			
		else
			if request("tfx_FldSort")="" then
				SQL="Insert Into " & RSCode("CodeTblName") & " (" & RSCode("CodeSrcFld") & "," & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & "," & RSCode("CodeSortFld") & ") Values(N'" & RSCode("CodeSrcItem") & "',N'" & request("tfx_value") & "',N'" &  request("tfx_display") & "',NULL)"			
			else
				'判斷為"國家與地區"代碼時，新增"地區"至mref欄位	
				if session("codeId") = "country_edit" then
					SQL="Insert Into " & RSCode("CodeTblName") & " (" & RSCode("CodeSrcFld") & "," & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & "," & RSCode("CodeSortFld") & ", mref) Values(N'" & RSCode("CodeSrcItem") & "',N'" & request("tfx_value") & "',N'" &  request("tfx_display") & "',N'" & request("tfx_FldSort") & "',N'" &  request("tfx_mref") & "')"		
				else
					SQL="Insert Into " & RSCode("CodeTblName") & " (" & RSCode("CodeSrcFld") & "," & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & "," & RSCode("CodeSortFld") & ") Values(N'" & RSCode("CodeSrcItem") & "',N'" & request("tfx_value") & "',N'" &  request("tfx_display") & "',N'" & request("tfx_FldSort") & "')"		
				end if
			end if
		end if			
	end if
'response.write SQL	
	conn.execute(SQL)%>
	<script language=vbs>
		alert("新增完成！")
		window.navigate "CodeDataDetailList.asp?codeId=<%=session("codeId")%>&CodeName=<%=session("CodeName")%>"
	</script>
<%	response.end
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
	
<% end sub '---- initForm() ----%>

<% sub showForm() %>

<!--#include file="CodeDataForm.inc"-->

<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
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
    <td class="Formtext" width="60%">【新增--代碼ID:<%=session("codeId")%>/代碼名稱:<%=session("CodeName")%>】</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><a href="Javascript:window.history.back();">回前頁</a>&nbsp;</td>   
	</td>
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
<script language="vbscript">

Sub formAdd
msg4 = "請務必輸入「Value值」，不得為空白！"
msg5 = "請務必輸入「Display值」，不得為空白！"

  If reg.tfx_Value.value = Empty Then
     MsgBox msg4, 64, "Sorry!"
     reg.tfx_Value.focus
     Exit Sub
  End if
  If reg.tfx_Display.value = Empty Then
     MsgBox msg5, 64, "Sorry!"
     reg.tfx_Display.focus
     Exit Sub
  End if  
  If bLen(reg.tfx_Value.value) > 20 Then
 	chky=msgbox("注意！"& vbcrlf & vbcrlf &"　代碼值超過8個字元,你確定新增資料嗎？"& vbcrlf , 48+1, "請注意！！")
 	if chky=vbok then  
  		reg.submitTask.value = "新增存檔"
  		reg.Submit   
     	end if
  else
  	reg.submitTask.value = "新增存檔"
  	reg.Submit        	
  end if  	     
End Sub

sub formReset
	reg.reset
end sub
</script>
<% end sub '--- showHTMLTail() ------%>
