<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="lightbluetable" align="right"><font color="red">*</font>
課程類別：</TD>
<TD class="whitetablebg"><Select name="htx_actCat" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID='ppActCat' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>
課程名稱：</TD>
<TD class="whitetablebg"><input name="htx_actName" size="40">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>
課程內容：</TD>
<TD class="whitetablebg"><textarea name="htx_actDesc" rows="6" cols="60">
</textarea>
</TD>
</TR>
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>
詳細資料：</TD>
<TD class="whitetablebg"><textarea name="htx_detailDesc" rows="6" cols="60">
</textarea>
</TD>
</TR>
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>
地點：</TD>
<TD class="whitetablebg"><input name="htx_place" size="50">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>
時間：</TD>
<TD class="whitetablebg"><input name="htx_actTime" size="50">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>
對象：</TD>
<TD class="whitetablebg"><input name="htx_actTarget" size="50">
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
		