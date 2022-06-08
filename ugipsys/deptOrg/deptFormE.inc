<form id="Form1" method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="deptEdit.asp?deptID=<% =request.querystring("deptID") %>">
<INPUT TYPE=hidden name=submitTask value="">
<input type="hidden" name="htx_deptID">
<%' ===begin========== Modified by Chris, 2006/08/24, to handle 單位組織作為CuDtSpecific ==================='%>
<input type="hidden" name="htx_giCuItem">
<%' ===end========== Modified by Chris, 2006/08/24, to handle 單位組織作為CuDtSpecific ==================='%>
<CENTER><TABLE  cellspacing="0">
<TR><TD class="Label" align="right"><font color="red">*</font>部門中文名稱：</TD>
<TD class="whitetablebg"><input name="htx_deptName" size="70">
</TD></TR>
<TR><TD class="Label" align="right">單位簡稱：</TD>
<TD class="whitetablebg"><input name="htx_AbbrName" size="30">
</TD></TR>
<TR><TD class="Label" align="right">英文名稱：</TD>
<TD class="whitetablebg"><input name="htx_eDeptName" size="50">
</TD></TR>
<TR><TD class="Label" align="right">英文簡稱：</TD>
<TD class="whitetablebg"><input name="htx_eAbbrName" size="30">
</TD></TR>
<%  	if checkGIPconfig("dpetOrgCode") then	%>
<TR><TD class="Label" align="right">單位代碼：</TD>
<TD class="whitetablebg"><input name="htx_deptCode" size="10">
<input name="htx_codeName" type="hidden">
</TD></TR>
<% 	end if %>
<%  	if checkGIPconfig("DeptHeadSet") then	%>
<TR><TD class="Label" align="right">單位主管：</TD>
<TD class="whitetablebg"><input name="htx_deptHead" size="10">
<input name="htx_DeptHeadUser" type="readonly" class="rdonly" size="20">
</TD></TR>
<% 	end if %>
<TR><TD class="Label" align="right">顯示次序：</TD>
<TD class="whitetablebg"><input name="htx_seq" size="2">
</TD></TR>
<TR><TD class="Label" align="right">單位住址：</TD>
<TD class="whitetablebg"><input name="htx_add" size="70">
</TD></TR>
<TR><TD class="Label" align="right">單位電話：</TD>
<TD class="whitetablebg"><input name="htx_tel" size="20">
</TD></TR>
<TR><TD class="Label" align="right">單位網址：</TD>
<TD class="whitetablebg"><input name="htx_website" size="70">
</TD></TR>
<TR><TD class="Label" align="right">單位層級：</TD>
<TD class="whitetablebg"><Select name="htx_OrgRank" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId=N'OrgRank' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select></TD></TR>
<TR><TD class="Label" align="right">類別：</TD>
<TD class="whitetablebg"><Select name="htx_kind" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId=N'OrgKind' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select></TD></TR>
<TR><TD class="Label" align="right">是否有效：</TD>
<TD class="whitetablebg"><Select name="htx_inUse" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId=N'boolYN' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select></TD></TR>
<% 	if checkGIPconfig("DeptTopCat") then %>
<TR><TD class="Label" align="right">資料大類：</TD>
<TD class="whitetablebg"><table width="100%"><TR>
			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId=N'topDataCat' Order by msortValue"
			SET RSS=conn.execute(SQL)
			pxCount = 0
			While not RSS.EOF
				if (pxCount mod 4) = 0 and pxCount<>0 then	response.write "</TR><TR>"
				pxCount = pxCount + 1
			%>
			<td><input type="checkbox" name="htx_tDataCat" value="<%=RSS(0)%>" ><font size=2><%=RSS(1)%></font></td>
			<%	RSS.movenext
			wend%>
</TR></table>
</TD></TR>
<%	end if %>
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
		<% if (RSreg("child") = 0) AND  (request.querystring("deptID") <> "0") then %>
            <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
		<% end if %>
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
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function
    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub
	
	sub htx_deptHead_onChange
		set oXML = createObject("Microsoft.XMLDOM")
		oXML.async = false
		xURI = "<%=session("mySysSiteURL")%>/wsxd/getUserName.asp?userID=" & reg.htx_deptHead.value
		oXML.load(xURI)	
		reg.htx_deptHeadUser.value = nullText(oXML.selectSingleNode("//mValue"))
	end sub
</SCRIPT>
	

<SCRIPT LANGUAGE="vbs">

</SCRIPT>
		