<%@ CodePage = 65001 %>
<%response.expires=0
HTProgCap="代碼維護"
HTProgCode="Pn90M02"
HTProgPrefix="CodeData" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
taskLable="編輯" & HTProgCap
	SQL="Select * from CodeMetaDef where codeId='" & session("codeID") & "'"
	SET RSCode=conn.execute(SQL)
if request("submittask")="編修存檔" then
  if request("tfx_Value")<>request("orgValue") then
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
<% 	response.end
   end if
  end if	
	if isNull(RSCode("CodeSrcFld")) then
	   if isNull(RSCode("CodeSortFld")) then
		SQL="Update " & RSCode("CodeTblName") & " SET " & RSCode("CodeValueFld") & "=N'" & request("tfx_value") & "'," & RSCode("CodeDisplayFld") & "=N'" & request("tfx_display") & "' where " & RSCode("CodeValueFld") & "=N'" & request("orgValue") & "' AND " & RSCode("CodeDisplayFld") & "=N'" & request("orgDisplay") & "'"	   
	   else	
		if request("tfx_FldSort")<>"" then	
			SQL="Update " & RSCode("CodeTblName") & " SET " & RSCode("CodeValueFld") & "=N'" & request("tfx_value") & "'," & RSCode("CodeDisplayFld") & "=N'" & request("tfx_display") & "'," & RSCode("CodeSortFld") & "=N'" & request("tfx_FldSort") & "' where " & RSCode("CodeValueFld") & "=N'" & request("orgValue") & "' AND " & RSCode("CodeDisplayFld") & "=N'" & request("orgDisplay") & "'"
		else
			SQL="Update " & RSCode("CodeTblName") & " SET " & RSCode("CodeValueFld") & "=N'" & request("tfx_value") & "'," & RSCode("CodeDisplayFld") & "=N'" & request("tfx_display") & "'," & RSCode("CodeSortFld") & "=NULL where " & RSCode("CodeValueFld") & "=N'" & request("orgValue") & "' AND " & RSCode("CodeDisplayFld") & "=N'" & request("orgDisplay") & "'"	
		end if		
	   end if
	else
	   if isNull(RSCode("CodeSortFld")) then
		SQL="Update " & RSCode("CodeTblName") & " SET " & RSCode("CodeValueFld") & "=N'" & request("tfx_value") & "'," & RSCode("CodeDisplayFld") & "=N'" & request("tfx_display") & "' where " & RSCode("CodeSrcFld") & "=N'" & RSCode("CodeSrcItem") & "' AND " & RSCode("CodeValueFld") & "=N'" & request("orgValue") & "' AND " & RSCode("CodeDisplayFld") & "=N'" & request("orgDisplay") & "'"	   	
	   else
		if request("tfx_FldSort")<>"" then
			'判斷為"國家與地區"代碼時，新增"地區"至mref欄位	
			if session("codeId") = "country_edit" then
				SQL="Update " & RSCode("CodeTblName") & " SET " & RSCode("CodeValueFld") & "=N'" & request("tfx_value") & "'," & RSCode("CodeDisplayFld") & "=N'" & request("tfx_display") & "'," & RSCode("CodeSortFld") & "=N'" & request("tfx_FldSort") & "', mref=N'" & request("tfx_mref") & "' where " & RSCode("CodeSrcFld") & "=N'" & RSCode("CodeSrcItem") & "' AND " & RSCode("CodeValueFld") & "=N'" & request("orgValue") & "' AND " & RSCode("CodeDisplayFld") & "=N'" & request("orgDisplay") & "'"
			else
				SQL="Update " & RSCode("CodeTblName") & " SET " & RSCode("CodeValueFld") & "=N'" & request("tfx_value") & "'," & RSCode("CodeDisplayFld") & "=N'" & request("tfx_display") & "'," & RSCode("CodeSortFld") & "=N'" & request("tfx_FldSort") & "' where " & RSCode("CodeSrcFld") & "=N'" & RSCode("CodeSrcItem") & "' AND " & RSCode("CodeValueFld") & "=N'" & request("orgValue") & "' AND " & RSCode("CodeDisplayFld") & "=N'" & request("orgDisplay") & "'"
			end if
		else
			SQL="Update " & RSCode("CodeTblName") & " SET " & RSCode("CodeValueFld") & "=N'" & request("tfx_value") & "'," & RSCode("CodeDisplayFld") & "=N'" & request("tfx_display") & "'," & RSCode("CodeSortFld") & "=NULL where " & RSCode("CodeSrcFld") & "=N'" & RSCode("CodeSrcItem") & "' AND " & RSCode("CodeValueFld") & "=N'" & request("orgValue") & "' AND " & RSCode("CodeDisplayFld") & "=N'" & request("orgDisplay") & "'"	
		end if
	   end if
	end if		
'response.write SQL			
	conn.execute(SQL)%>
	<script language=vbs>
		alert("編修完成！")
		window.navigate "CodeDataDetailList.asp?codeId=<%=session("codeId")%>&CodeName=<%=session("CodeName")%>"
	</script>	
<%
	response.end	
elseif request("submittask")="刪除" then
'	SQL="Select * from CodeTable where codeId='" & session("codeID") & "'"	
'	SET RSCode=conn.execute(SQL)
	if isNull(RSCode("CodeSrcFld")) then
		SQL="Delete from " & RSCode("CodeTblName") & " where " & RSCode("CodeValueFld") & "=N'" & request("orgValue") & "' AND " & RSCode("CodeDisplayFld") & "=N'" & request("orgDisplay") & "'"	
	else	
		SQL="Delete from " & RSCode("CodeTblName") & " where " & RSCode("CodeSrcFld") & "=N'" & RSCode("CodeSrcItem") & "' AND " & RSCode("CodeValueFld") & "=N'" & request("orgValue") & "' AND " & RSCode("CodeDisplayFld") & "=N'" & request("orgDisplay") & "'"
	end if		
	'response.write sql
	conn.execute(SQL)%>
	<script language=vbs>
		alert("刪除完成！")
		window.navigate "CodeDataDetailList.asp?codeId=<%=session("codeId")%>&CodeName=<%=session("CodeName")%>"
	</script>	
<%	response.end
elseif 1=2 then		
'---- Modified by Chris 2006/4/6 -------- begin -------------------------------------------------------------------
'---- 	put the actual else part back in order to get data from tables other than CodeMain -------------------
	xvalue=request.querystring("value")
        sql="select * from CodeMain where codeMetaId='" & request("codeId") & "' and mcode='" & xvalue & "'"
    'response.write sql & "<HR>"
set rs=Conn.Execute(sql)
if not rs.eof then
	xDisplay=rs("mvalue")
	xFldSort=rs("msortValue")
	if session("codeId") = "country_edit" then
	xmref=rs("mref")
	end if
end if
	showHTMLHead()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
else
   if isNull(RSCode("CodeSrcFld")) then
	SQL="Select " & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld")  & "," & RSCode("CodeSortFld") & " from " & RSCode("CodeTblName") & " where " & RSCode("CodeValueFld") & "='" & request.querystring("value") & "'"
   else
	SQL="Select " & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld")  & "," & RSCode("CodeSortFld") & " from " & RSCode("CodeTblName") & " where " & RSCode("CodeSrcFld") & "='" & RSCode("CodeSrcItem") & "' AND " & RSCode("CodeValueFld") & "='" & request.querystring("value") & "'"
   end if
   	Set RSreg = conn.execute(SQL)
	xvalue=request.querystring("value")
	if not RSreg.eof then
		xDisplay=RSreg(1)
		xFldSort=RSreg(2)
		if session("codeId") = "country_edit" then	xmref=RSreg("mref")
	end if
	showHTMLHead()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
'---- Modified by Chris 2006/4/6 -------------- end --------------------------------------------------------------
'---- 	put the actual else part back in order to get data from tables other than CodeMain -------------------
end if	
%>
<% sub initForm() %>
	
<% end sub '---- initForm() ----%>

<% sub showForm() %>

<!--#include file="CodeDataForm.inc"-->

<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Query.asp">
<title>查詢表單</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
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
    <td class="Formtext" width="60%">【編輯--代碼ID:<%=session("codeId")%>/代碼名稱:<%=session("CodeName")%>】</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><a href="Javascript:window.history.back();">回前頁</a>&nbsp;</td>
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
sub window_onLoad
	Initform
end sub

sub Initform
	reg.orgValue.value="<%=xvalue%>"
	reg.orgDisplay.value="<%=xdisplay%>"
	reg.tfx_Value.value="<%=xvalue%>"
	reg.tfx_Display.value="<%=xdisplay%>"
<%if session("codeId") = "country_edit" then%>
	reg.tfx_mref.value="<%=xmref%>"
<%end if%>	
<%if session("SortFlag") then%>	
	reg.orgFldSort.value="<%=xFldsort%>"
	reg.tfx_FldSort.value="<%=xFldsort%>"
<%end if%>	
end sub

Sub formEdit
msg4 = "請務必輸入「Value值」，不得為空白！"
msg5 = "請務必輸入「Display值」，不得為空白！"
msg6 = "請務必輸入「過濾(排序)值」，不得為空白！"

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
 	chky=msgbox("注意！"& vbcrlf & vbcrlf &"　代碼值超過8個字元,你確定編修資料嗎？"& vbcrlf , 48+1, "請注意！！")
 	if chky=vbok then  
  		reg.submitTask.value = "編修存檔"
  		reg.Submit   
     	end if
  else
  	reg.submitTask.value = "編修存檔"
  	reg.Submit        	
  end if    	     
End Sub

Sub formDelete
 chky=msgbox("注意！"& vbcrlf & vbcrlf &"　你確定刪除資料嗎？"& vbcrlf , 48+1, "請注意！！")
 if chky=vbok then
	reg.submitTask.value = "刪除"
	reg.Submit
 End If
End Sub

sub formReset
	reg.reset
	Initform
end sub
</script>
<% end sub '--- showHTMLTail() ------%>
