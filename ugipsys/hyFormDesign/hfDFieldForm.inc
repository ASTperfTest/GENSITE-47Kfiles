<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=org_fieldName value="">
<CENTER>
<table cellspacing=2 border=1>
<TR><TH>標題</TH>
<TD><INPUT name="xuf_fieldLabel" size="30"></TD>
<TR><TH>欄位名稱</TH>
<TD><INPUT name="xuf_fieldName" size="20"></TD>
</TR>
<TR><TH>資料型別</TH>
<TD><SELECT name="xuf_dataType" size="1">
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'htDdataType' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(1)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select></TD>
</TR>
<TR><TH>輸入方式</TH>
<TD><SELECT name="xuf_inputType" size="1">
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'htDinputType' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select></TD>
</TR>
<TR><TH>長度</TH>
<TD><INPUT name="xuf_inputLen" size="2"></TD>
</TR>
<TR><TH>非必填</TH>
<TD>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'boolYN' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

<input type="radio" name="xuf_canNull" value="<%=RSS(0)%>" <%=pdxc%>>
			<%=RSS(1)%>&nbsp;&nbsp;
			<%	RSS.movenext
			wend%>
</TR>
<TR><TH>選項條列</TH>
<TD><TEXTAREA name="xuf_itemList" rows="5" cols="40"></TEXTAREA></TD>
</TR>
<% if request("part")="formList" then %>
<TR><TH>標題HTML</TH>
<TD><TEXTAREA name="xuf_fieldDesc" rows="5" cols="60"></TEXTAREA></TD>
</TR>
<% else %>
<TR><TH>表頭欄位</TH>
<TD><SELECT name="xuf_isPrimaryKey" size="1">
<OPTION value="Y">否</OPTION>
<OPTION value="P">是</OPTION>
</SELECT></TD></TR>
<% end if %>
</TABLE>
</form>     
</CENTER>
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
		