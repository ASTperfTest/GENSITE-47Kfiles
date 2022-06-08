<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="eTableLable" align="right"><font color="red">*</font>
附件名稱：</TD>
<TD class="eTableContent"><input name="htx_aTitle" size="50">
<input type="hidden" name="htx_ixCuAttach">
<input type="hidden" name="htx_xiCuItem">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">說明：</TD>
<TD class="eTableContent"><input name="htx_aDesc" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
系統檔名：</TD>
<TD class="eTableContent"><table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr><td colspan="2">
		<input type="file" name="htFile_NFileName">
		<input type="hidden" name="hoFile_NFileName">
		<input type="hidden" name="htFileActCK_NFileName">
	</td></tr>
	<tr>
	<td><span id="logo_NFileName" class="rdonly"></span>
		<div id="noLogo_NFileName" style="color:red">無附檔</div></td>
	<td valign="bottom">
		<div id="LbtnHide0_NFileName">
		<% if (HTProgRight and 4)<>0 then %>
			<img id="addLogo_NFileName" class="hand" src="../pagestyle/addimg.gif" onClick="VBS: addXFile 'NFileName'">
		<% End if %>
		</div>
		<div id="LbtnHide1_NFileName">
		<% if (HTProgRight and 8)<>0 then %>
			<img id="chgLogo_NFileName" class="hand" src="../pagestyle/chimg.gif" onClick="VBS: chgXFile 'NFileName'">
		<% End if %>
		<% if (HTProgRight and 16)<>0 then %>
			<img id="delLogo_NFileName" class="hand" src="../pagestyle/delimg.gif" onClick="VBS: delXFile 'NFileName'">
		<% End if %>
		</div>
		<div id="LbtnHide2_NFileName">
		<% if (HTProgRight and 8)<>0 then %>
			<img id="orgLogo_NFileName" class="hand" src="../pagestyle/resetimg.gif" onClick="VBS: orgXFile 'NFileName'">
		<% End if %>
		</div>
	</td></tr></table>
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
顯示於附件條列：</TD>
<TD class="eTableContent"><Select name="htx_bList" size=1>
<option value="Y">是</option>
<option value="N">否</option>
</select>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">條列次序：</TD>
<TD class="eTableContent"><input name="htx_listSeq" size="2">
</TD>
</TR>
</TABLE>
</CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
<%If formFunction = "query" then %>
        <%if (HTProgRight AND 2) <> 0 then %>
            <input type=button value ="查　詢" class="cbutton" onClick="formSearchSubmit()">
            <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <%end if%>    
<%Elseif formFunction = "edit" then %>
        <%if (HTProgRight AND 8) <> 0 then %>
               <input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()">
        <% End If %>
        
       <% if (HTProgRight AND 16) <> 0 then %>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>           
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">

<% Else '-- add ---%>          
          <%if (HTProgRight AND 4)<>0 then %>
               <input type=button value ="新增存檔" class="cbutton" onClick="formAddSubmit()">
          <%end if%>     
          
          <%if (HTProgRight AND 4)<>0  then %>
              <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
          <%end if%>    

<% End If %>
 </td></tr>
</table>
<SCRIPT LANGUAGE="vbs">
    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub
</SCRIPT>
	

<SCRIPT LANGUAGE="vbs">

</SCRIPT>
		