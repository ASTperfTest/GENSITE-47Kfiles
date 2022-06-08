<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="80%"><TR><TD class="lightbluetable" align="right"><font color="red">*</font>
資料欄位名稱：</TD>
<TD class="whitetablebg" colspan="3"><input name="htx_xfieldName" size="30" class="rdonly">
<input type="hidden" name="htx_entityID">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">標題：</TD>
<TD class="whitetablebg" colspan="3"><input name="htx_xfieldLabel" size="30">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">說明：</TD>
<TD class="whitetablebg" colspan="3"><input name="htx_xfieldDesc" size="50">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">排序值：</TD>
<TD class="whitetablebg" colspan="3"><input name="htx_xfieldSeq" size=4>
</TD>
</TR>
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>
資料型別：</TD>
<TD class="whitetablebg" colspan="3"><Select name="htx_xdataType" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'htDdataType' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="lightbluetable" align="right">資料長度：</TD>
<TD class="whitetablebg"><input name="htx_xdataLen" size="4">
</TD>
<TD class="lightbluetable" align="right">輸入長度：</TD>
<TD class="whitetablebg"><input name="htx_xinputLen" size="4">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">允許空值：</TD>
<TD class="whitetablebg">			<% pdxc = ""%>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'boolYN' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

<input type="radio" name="htx_xcanNull" value="<%=RSS(0)%>" <%=pdxc%>>
			<%=RSS(1)%>&nbsp;&nbsp;
			<%	RSS.movenext
			wend%>
</TD>
<TD class="lightbluetable" align="right">為鍵值：</TD>
<TD class="whitetablebg">			<% pdxc = ""%>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'boolYN' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

<input type="radio" name="htx_xisPrimaryKey" value="<%=RSS(0)%>" <%=pdxc%>>
			<%=RSS(1)%>&nbsp;&nbsp;
			<%	RSS.movenext
			wend%>
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">identity：</TD>
<TD class="whitetablebg">			<% pdxc = ""%>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'boolYN' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

<input type="radio" name="htx_xidentity" value="<%=RSS(0)%>" <%=pdxc%>>
			<%=RSS(1)%>&nbsp;&nbsp;
			<%	RSS.movenext
			wend%>
</TD>
<TD class="lightbluetable" align="right">dbDefault：</TD>
<TD class="whitetablebg"><input name="htx_xdefaultvalue" size="20">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">clientDefault：</TD>
<TD class="whitetablebg" colspan="3"><input name="htx_xclientDefault" size="50">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">輸入型式：</TD>
<TD class="whitetablebg"><Select name="htx_xinputType" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID=N'htDinputType' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
<TD class="lightbluetable" align="right">xrefLookup：</TD>
<TD class="whitetablebg"><input name="htx_xrefLookup" size=20>
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">rows：</TD>
<TD class="whitetablebg"><input name="htx_xrows" size=4>
</TD>
<TD class="lightbluetable" align="right">cols：</TD>
<TD class="whitetablebg"><input name="htx_xcols" size=4>
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
		