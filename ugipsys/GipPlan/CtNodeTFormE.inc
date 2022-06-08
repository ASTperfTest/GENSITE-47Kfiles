<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="eTableLable" align="right"><font color="red">*</font>
節點類別：</TD>
<TD class="eTableContent"><Select name="htx_CtNodeKind" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from codeMain where mSortValue IS NOT NULL AND codeMetaID=N'refCtNodeKind' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

<input type="hidden" name="htx_CtRootID">
<input type="hidden" name="htx_EditUserID">
<input type="hidden" name="htx_EditDate">
<input type="hidden" name="htx_CtNodeID">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
標題：</TD>
<TD class="eTableContent"><input name="htx_CatName" size="30">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
主題單元：</TD>
<TD class="eTableContent"><Select name="htx_CtUnitID" size=1>
<option value="">請選擇</option>
			<%SQL="Select CtUnitID,CtUnitName from CtUnit where CtUnitID IS NOT NULL AND inUse='Y' Order by CtUnitID"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right">標題圖示：</TD>
<TD class="eTableContent"><table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr><td colspan="2">
		<input type="file" name="htImg_CtNameLogo">
		<input type="hidden" name="hoImg_CtNameLogo">
		<input type="hidden" name="htImgActCK_CtNameLogo">
	</td></tr>
	<tr>
	<td width="37%"><img id="logo_CtNameLogo" src="">
		<div id="noLogo_CtNameLogo" style="color:red">無圖片</div></td>
	<td valign="bottom">
		<div id="LbtnHide0_CtNameLogo">
		<% if (HTProgRight and 4)<>0 then %>
			<img id="addLogo_CtNameLogo" class="hand" src="../pagestyle/addimg.gif" onClick="VBS: addLogo 'CtNameLogo'">
		<% End if %>
		</div>
		<div id="LbtnHide1_CtNameLogo">
		<% if (HTProgRight and 8)<>0 then %>
			<img id="chgLogo_CtNameLogo" class="hand" src="../pagestyle/chimg.gif" onClick="VBS: chgLogo 'CtNameLogo'">
		<% End if %>
		<% if (HTProgRight and 16)<>0 then %>
			<img id="delLogo_CtNameLogo" class="hand" src="../pagestyle/delimg.gif" onClick="VBS: delLogo 'CtNameLogo'">
		<% End if %>
		</div>
		<div id="LbtnHide2_CtNameLogo">
		<% if (HTProgRight and 8)<>0 then %>
			<img id="orgLogo_CtNameLogo" class="hand" src="../pagestyle/resetimg.gif" onClick="VBS: orgLogo 'CtNameLogo'">
		<% End if %>
		</div>
	</td></tr></table>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">顯示次序：</TD>
<TD class="eTableContent"><input name="htx_CatShowOrder" size="1">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
是否開放：</TD>
<TD class="eTableContent"><Select name="htx_inUse" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from codeMain where mSortValue IS NOT NULL AND codeMetaID=N'bOnOff' Order by mSortValue"
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
		