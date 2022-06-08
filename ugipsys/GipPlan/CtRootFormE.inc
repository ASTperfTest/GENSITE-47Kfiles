<form id="Form1" method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="CtRootEdit.asp?ctRootId=<% =request.querystring("ctRootId") %>">
<INPUT TYPE=hidden name=submitTask value="">
<table cellspacing="0">
<TR><TD align="right" class="Label"><font color="red">*</font>
目錄樹名稱：</TD>
<TD class="eTableContent"><input name="htx_ctRootName" size="40">
<input type="hidden" name="htx_EditDate">
<input type="hidden" name="htx_Editor">
<input type="hidden" name="htx_CtRootID">
</TD>
</TR>
<TR><TD align="right" class="Label"><font color="red">*</font>
類別：</TD>
<TD class="eTableContent"><Select name="htx_vGroup" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId=N'visitGroup' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD align="right" class="Label">用途說明：</TD>
<TD class="eTableContent"><input name="htx_Purpose" size="50">
</TD>
</TR>
<TR><TD align="right" class="Label"><font color="red">*</font>
是否開放：</TD>
<TD class="eTableContent"><Select name="htx_inUse" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId=N'bOnOff' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
</TD></TR>
<TR><TD align="right" class="Label">單 位：</TD>
<TD class="eTableContent"><Select name="htx_deptId" size=1>
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
</TD></TR>
<TR><TD align="right" class="Label">上稿樹/預覽版型：</TD>
<TD class="eTableContent"><Select name="htx_pvXdmp" size=1>
<option value="">--無--</option>
			<%SQL="Select xdmpId,xdmpName from XdmpList Order by xdmpId"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
  <font size="-1">　(選 --無-- 則表示不作上稿樹)</font>
</TD></TR>
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
		
