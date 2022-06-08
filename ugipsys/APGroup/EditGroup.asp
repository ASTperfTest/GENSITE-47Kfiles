<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "HT001" 
' ============= Modified by Chris, 2006/09/12, to handle 上稿權限能以使用者設定 聯集 群組設定 ========================'
'		Document: 950912_智庫GIP擴充.doc
'  modified list:
'	加 include /inc/checkGIPconfig.inc	'-- 判別是否用 <CtUgrpSet>
'	加判別 <CtUgrpSet> 上方產生『上稿權限』連結
'		先只產生上稿權限 (aType=A)，審稿或複審需要時可加連結用不同參數
%>
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
if request("ck") = "" then
   ugrpID = Request.Form("ugrpID")
   sqlcom = "select * from Ugrp Where ugrpId =N'"& ugrpID & "'"
   Set RS = Conn.execute(sqlcom)
elseif request("ck") = "Edit" Then
        ugrpID = request("ugrpID")
		sqlCom = "Update Ugrp Set "
		sqlCom = sqlCom & "ugrpName=" & pkStr(request("ugrpName"),",")
		sqlCom = sqlCom & "remark=" & pkStr(request("remark"),",")
		sqlCom = sqlCom & "isPublic=" & pkStr(request("isPublic"),",")
		sqlCom = sqlCom & "regdate=N'" & Request("RegDate") & "' Where ugrpId =N'"& ugrpID & "'"
		Conn.execute(sqlCom)
%>
			 <script language=VBScript>
	  			alert("已成功存檔！")
	  			document.location.href="ListGroup.asp"
			 </script>
<%  Response.End
        
ElseIF request("ck") = "Del" then	
   ugrpID = request("ugrpID")	
	SQL1 = "Update InfoUser Set ugrpId = N'basic' where ugrpId=N'" & ugrpID & "'"
	set rs1 = conn.Execute(SQL1) 

	SQL = "Delete from Ugrp Where ugrpId =N'"& ugrpID & "'"
	set rs = conn.Execute(SQL) 

	SQL = "Delete from UgrpAp Where ugrpId =N'"& ugrpID & "'"
	set rs = conn.Execute(SQL) 
	
%>
	 <script language=VBScript>
			alert("已成功刪除！所有成員也成功轉【基本群組】！")
			document.location.href="ListGroup.asp"
	 </script>
<%  Response.End

end if %>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>

<body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
<%
' ===begin========== Modified by Chris, 2006/09/12, to handle 上稿權限能以使用者設定 聯集 群組設定 ========================
   if checkGIPconfig("CtUgrpSet") then %>
	<a href="CtUgrpSet.asp?aType=A&ugrpID=<%=RS("ugrpID")%>">上稿權限</a>
<% end if 
' ===end========== Modified by Chris, 2006/09/12, to handle 上稿權限能以使用者設定 聯集 群組設定 ========================
%>
    <% if (HTProgRight and 4)=4 then %><a href="Addgroup.asp">新增</a><% End IF %>
	<% if (HTProgRight and 2)=2 then %>
		<a href="Querygroup.asp">查詢</a>
	    <a href="ListUser.asp?ugrpID=<%=RS("ugrpID")%>">檢視成員</a>
	    <a href="VBScript: history.back">回上一頁</a>
	<% End IF %>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">編修權限群組</div>

<form id="Form1" name=reg method="POST">
    <input name="ck" type="hidden" value="">
    <input name=RegDate type="hidden" value="<%=RS("RegDate")%>">    
  <table cellspacing="0">
     <tr>    
      <td align="right" class="Label">群組ID</td>     
      <td class="whitetablebg"><input name="ugrpID" size="20" value="<%=rs("ugrpID")%>" readonly class=sedit> </td>     
     </tr>
     <tr>    
      <td align="right" class="Label"><span class="Must">*</span>群組名稱</td>     
      <td class="whitetablebg"><input name="ugrpName" size="30" value="<%=rs("ugrpName")%>"> </td>     
     </tr>     
     <tr>       
      <td align="right" class="Label">說明</td>      
      <td colspan=3 class="whitetablebg"><input name="remark" size="30" value="<%=rs("remark")%>"></td>      
    </tr>
	<TR><TD class="Label" align="right">是否公開：</TD>
	<TD class="whitetablebg"><Select name="isPublic" size=1>
	<option value="">請選擇</option>
				<%SQL="Select mcode,mvalue from CodeMain where codeMetaId=N'boolYN' Order by msortValue"
				SET RSS=conn.execute(SQL)
				While not RSS.EOF%>
	
				<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
				<%	RSS.movenext
				wend%>
	</select></TD></TR>
     <tr>    
      <td align="right" class="Label">建檔日期</td>     
      <td class="whitetablebg"><input name="regdateShow" size="20" readonly class=sedit value="<%=d7date(RS("RegDate"))%>"> </td>     
     </tr>  
     <tr>    
      <td align="right" class="Label">資料建檔人</td>     
      <td class="whitetablebg"><input name="signer" size="20" value="<%=RS("signer")%>" class=sedit readonly> </td>     
     </tr>                
	</table>      
   <% if (HTProgRight and 8)=8 then %><input type="button" value="編修存檔" name="Enter" class="cbutton" OnClick="datacheck()"><% End IF %> 
   <input type=reset class=cbutton value="清除重填">
   <% if (HTProgRight and 16)=16 and RS("ugrpID")<>"basic" then %><input type="button" value="刪除群組" name="Enter" class="cbutton" OnClick="del()"><% End If %>   
</form>
   
<div id="Explain">
	<h1>說明</h1>
	<ul>
		<li><span class="Must">*</span>為必要欄位</li>
		<li>不設定任何條件可查詢全部權限群組資料</li> 
	</ul>
</div>
</body>
</html>        
<script language="vbscript">
msg1 = "「群組名稱」不得空白！"

reg.isPublic.value = "<%=RS("isPublic")%>"

Sub datacheck()    
  If reg.ugrpName.value = Empty Then
     MsgBox msg1, 64, "Sorry!"
     reg.ugrpName.focus
     Exit Sub
  End if 
  reg.ck.value="Edit"
  reg.Submit
End Sub

Sub del()
 chky=msgbox("注意！"& vbcrlf & vbcrlf &"你確定刪除此群組資料嗎？"& vbcrlf & vbcrlf &"若是,此群組所有成員將轉為【基本群組】！"& vbcrlf & vbcrlf & "你確定嗎？", 48+1, "請注意！！")
 if chky=vbok then
   reg.ck.value="Del"
   reg.Submit
 End If
End Sub

</script>