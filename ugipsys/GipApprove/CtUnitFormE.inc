<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="eTableLable" align="right">主題單元名稱：</TD>
<TD class="eTableContent"><input name="htx_CtUnitName" size="30" title="作為分類節點的預設值">
<input type="hidden" name="htx_CtUnitID">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">單元LOGO：</TD>
<TD class="eTableContent"><table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr><td colspan="2">
		<input type="file" name="htImg_CtUnitLogo">
		<input type="hidden" name="hoImg_CtUnitLogo">
		<input type="hidden" name="htImgActCK_CtUnitLogo">
	</td></tr>
	<tr>
	<td width="37%"><img id="logo_CtUnitLogo" src="">
		<div id="noLogo_CtUnitLogo" style="color:red">無圖片</div></td>
	<td valign="bottom">
		<div id="LbtnHide0_CtUnitLogo">
		<% if (HTProgRight and 4)<>0 then %>
			<img id="addLogo_CtUnitLogo" class="hand" src="../pagestyle/addimg.gif" onClick="VBS: addLogo 'CtUnitLogo'">
		<% End if %>
		</div>
		<div id="LbtnHide1_CtUnitLogo">
		<% if (HTProgRight and 8)<>0 then %>
			<img id="chgLogo_CtUnitLogo" class="hand" src="../pagestyle/chimg.gif" onClick="VBS: chgLogo 'CtUnitLogo'">
		<% End if %>
		<% if (HTProgRight and 16)<>0 then %>
			<img id="delLogo_CtUnitLogo" class="hand" src="../pagestyle/delimg.gif" onClick="VBS: delLogo 'CtUnitLogo'">
		<% End if %>
		</div>
		<div id="LbtnHide2_CtUnitLogo">
		<% if (HTProgRight and 8)<>0 then %>
			<img id="orgLogo_CtUnitLogo" class="hand" src="../pagestyle/resetimg.gif" onClick="VBS: orgLogo 'CtUnitLogo'">
		<% End if %>
		</div>
	</td></tr></table>
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
單元類型：</TD>
<TD class="eTableContent"><Select name="htx_CtUnitKind" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from CodeMainLong where mSortValue IS NOT NULL AND codeMetaID=N'refCTUKind' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right">重導URL：</TD>
<TD class="eTableContent"><input name="htx_redirectURL" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">開新視窗：</TD>
<TD class="eTableContent"><Select name="htx_newWindow" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'boolYN' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right">單元資料定義：</TD>
<TD class="eTableContent"><Select name="htx_iBaseDSD" size=1>
<option value="">請選擇</option>
			<%SQL="Select iBaseDSD,sBaseDSDName from BaseDSD where iBaseDSD IS NOT NULL AND inUse='Y' Order by iBaseDSD"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
只顯示此單元資料：</TD>
<TD class="eTableContent"><Select name="htx_fCtUnitOnly" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'boolYN' Order by mSortValue"
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
		