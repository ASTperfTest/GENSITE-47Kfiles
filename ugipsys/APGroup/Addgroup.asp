<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "HT001"
%>
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual = "/inc/dbutil.inc" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
Function TranComp(groupcode,value)
   if (groupcode and value )=value then
      response.write "checked"
   end if
End function

if request("submitTask") = "新增存檔" then
	sqlCom = "SELECT * FROM Ugrp WHERE ugrpId = '" & request("fgugrpID") & "'"
	Set RS = Conn.execute(sqlCom)
	if not RS.EOF then %>
	<script language=VBScript>
	  alert("此群組ID已建立，無法再次建立！請另取ID！！")
	  window.history.back
	</script>	
<%	else
		sqlCom = "INSERT INTO Ugrp (ugrpId,ugrpName,remark,isPublic,regdate,signer) VALUES ("
		sqlCom = sqlCom & pkStr(request("fgugrpID"),",")		
		sqlCom = sqlCom & pkStr(request("fgugrpName"),",")
		sqlCom = sqlCom & pkStr(request("fgremark"),",")
		sqlCom = sqlCom & pkStr(request("isPublic"),",")
		sqlCom = sqlCom & "'" & request("fgregdate") & "',"
		sqlCom = sqlCom & "'" & session("userID") & "')"				
		Conn.execute(sqlCom)
		SQL="Insert Into UgrpAp (ugrpId,apcode,rights,regdate) Values(N'" & _
			request("fgugrpID") & "',N'Pn00M00',0,N'" & date() & "')"
		conn.execute(SQL)
		%>
		<script language=VBScript>
		  alert("新增群組基本資料完成，請設定群組之權限！")
		  window.navigate "AddRegGroup.asp?ugrpName=<%=request("fgugrpName")%>&ugrpID=<%=request("fgugrpID")%>"
		</script>	
<%     		response.end
	end if
end if
%>
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
 <% if (HTProgRight and 2)=2 then %><a href="Querygroup.asp">查詢</a><% End IF %>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">新增權限群組</div>

<form id="Form1" name=fmGrpAdd method="POST">
      <input type=hidden name=fgregdate value="<%=date()%>">
      <input type=hidden name=submitTask>
  <table cellspacing="0">
     <tr>    
      <td align="right" class="Label"><span class="Must">*</span>群組ID</td>     
      <td class="whitetablebg"><input name="fgugrpID" class="InputText" size="15"> </td>     
     </tr>
     <tr>    
      <td align="right" class="Label"><span class="Must">*</span>群組名稱</td>     
      <td class="whitetablebg"><input name="fgugrpName" class="InputText" size="15"> </td>     
     </tr>     
     <tr>       
      <td align="right" class="Label">說明</td>      
      <td class="whitetablebg"><input name="fgRemark" class="InputText" size="40"></td>      
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
      <td class="whitetablebg"><input name="fgregdateShow" size="15" readonly class=sedit value="<%=d7date(date())%>"> </td>     
     </tr>  
     <tr>    
      <td align="right" class="Label">資料建檔人</td>     
      <td class="whitetablebg"><input name="fgsigner" size="15" value="<%=session("UserID")%>" class=sedit readonly> </td>     
     </tr>  
      	</table>      
            <% if (HTProgRight and 4)=4 then %>
            	<input type=button value="新增存檔" class="cbutton" onclick="FormAddCheck">
            <% End If %>  
               <input type=reset class=cbutton value="清除重填">            
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
<script language=VBS>     
sub FormAddCheck    
	if fmGrpAdd.fgugrpID.value = "" then     
		alert "群組ID欄位不可空白"
		fmGrpAdd.fgugrpID.focus
		exit sub
	end if
	if fmGrpAdd.fgugrpName.value = "" then    
		alert "群組名稱欄位不可空白"
		fmGrpAdd.fgugrpName.focus		   
		exit sub     
	end if
	fmGrpAdd.submitTask.value="新增存檔"	
        fmGrpAdd.submit
end sub   
</script>     
