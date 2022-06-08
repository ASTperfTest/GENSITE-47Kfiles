<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="80%"><TR><TD class="lightbluetable" align="right"><font color="red">*</font>
資料源代碼：</TD>
<TD class="whitetablebg"><input name="htx_dbId" size="10" readonly="true" class="rdonly">
<input type="hidden" name="htx_entityID">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>
類別：</TD>
<TD class="whitetablebg"><Select name="htx_apcatID" size=1>
<option value="">請選擇</option>
			<%SQL="Select apcatId,apcatCname from Apcat Order by apseq"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
</TD>
</TR>
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>
資料表名稱：</TD>
<TD class="whitetablebg"><input name="htx_tableName" size="20">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">說明：</TD>
<TD class="whitetablebg"><input name="htx_entityDesc" size="50">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">資料來源URI：</TD>
<TD class="whitetablebg"><input name="htx_entityURI" size="50">
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
		