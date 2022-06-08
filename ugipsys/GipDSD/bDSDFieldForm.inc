<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="bDSDFieldAdd.asp">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="eTableLable" align="right">顯示次序：</TD>
<TD class="eTableContent"><input name="htx_xfieldSeq" size="2" readonly="true" class="rdonly">
<input type="hidden" name="htx_iBaseDSD">
<input type="hidden" name="htx_iBaseField">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
是否生效：</TD>
<TD class="eTableContent"><Select name="htx_inUse" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='boolYN' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right">標題：</TD>
<TD class="eTableContent"><input name="htx_xfieldLabel" size="30">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">說明：</TD>
<TD class="eTableContent"><input name="htx_xfieldDesc" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
資料型別：</TD>
<TD class="eTableContent"><Select name="htx_xdataType" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='htDdataType' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right">資料長度：</TD>
<TD class="eTableContent"><input name="htx_xdataLen" size="4">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
允許空值：</TD>
<TD class="eTableContent">			<% pdxc = ""%>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='boolYN' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

<input type="radio" name="htx_xcanNull" value="<%=RSS(0)%>" <%=pdxc%>>
			<%=RSS(1)%>&nbsp;&nbsp;
			<%	RSS.movenext
			wend%>
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
輸入型式：</TD>
<TD class="eTableContent"><Select name="htx_xinputType" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='htDinputType' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right">xrefLookup：</TD>
<TD class="eTableContent"><input name="htx_xrefLookup" size="20">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">欄位名稱：</TD>
<TD class="eTableContent"><input name="htx_xfieldName" size="30">
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
		