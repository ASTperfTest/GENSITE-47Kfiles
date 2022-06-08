<form id="Form1" method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="xdmpEdit.asp?xdmpID=<% =request.querystring("xdmpID") %>">
<INPUT TYPE=hidden name=submitTask value="">
<input type="hidden" name="htx_Editor">
<input type="hidden" name="htx_EditDate">
<table cellspacing="0">
     <tr>    
      <td align="right" class="Label"><span class="Must">*</span>版面ID</td>     
      <td class="whitetablebg"><input name="htx_xdmpID" class="rdonly" size="6" readonly> </td>     
     </tr>
     <tr>    
      <td align="right" class="Label"><span class="Must">*</span>版面名稱</td>     
      <td class="whitetablebg"><input name="htx_xdmpName" class="InputText" size="50"> </td>     
     </tr>
     <tr>    
      <td align="right" class="Label">用途說明</td>     
      <td class="whitetablebg"><input name="htx_Purpose" class="InputText" size="80"> </td>     
     </tr>
     <tr>    
      <td align="right" class="Label"><span class="Must">*</span>單 位</td>     
      <td class="whitetablebg"><Select name="htx_deptId" size=1>
<option value="">--- 共 用 ---</option>
<%
	sqlCom ="SELECT D.deptId,D.deptName,D.parent,len(D.deptId)-1,D.seq," & _
		"(Select Count(*) from Dept WHERE parent=D.deptId AND nodeKind='D') " & _
		"  FROM Dept AS D Where D.nodeKind='D' " _
		& " AND D.deptId LIKE N'" & session("deptId") & "%'" _
		& " ORDER BY len(D.deptId), D.parent, D.seq"	
'response.write SqlCom
	set RSS = conn.execute(sqlCom)
	if not RSS.EOF then
		ARYDept = RSS.getrows(300)
		glastmsglevel = 0
		genlist 0, 0, 1, 0
	        expandfrom ARYDept(cid, 0), 0, 0
	end if
%>
</select>
	</td></tr>
</table>
        <%if (HTProgRight AND 8) <> 0 then %>
               <input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()">
        <% End If %>
        
       <% if (HTProgRight AND 16) <> 0 then %>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>           
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">
</form>  
   
<div id="Explain">
	<h1>說明</h1>
	<ul>
		<li><span class="Must">*</span>為必要欄位</li>
	</ul>
</div>
<SCRIPT LANGUAGE="vbs">
    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub
</SCRIPT>
			
