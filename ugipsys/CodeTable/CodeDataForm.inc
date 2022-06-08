<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=orgValue value="">
<INPUT TYPE=hidden name=orgDisplay value="">
<%if session("SortFlag") then%>	
<INPUT TYPE=hidden name=orgFldSort value="">
<%end if%>
<CENTER><TABLE border=0 class=bluetable cellspacing=1 cellpadding=2 width="70%">
		<TR>
		  <TD class=lightbluetable align=right><%=RSCode("CodeValueFldName")%>(Value)</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_Value SIZE=15 <%IF request.querystring("CodeID") <> "" then %> class=sedit readonly<% End If %>></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right><%=RSCode("CodeDisplayFldName")%>(Display)</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_Display SIZE=30></TD>
		</TR>
		<%if session("SortFlag") then%>
		<TR>
		  <TD class=lightbluetable align=right><%=RSCode("CodeSortFldName")%>(SORT)</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_FldSort SIZE=15></TD>
		</TR>
		<%end if%>
		<% '判斷為國家與地區代碼時，新增地區欄位選項 %>
		<%if session("codeId") = "country_edit" then%>
		<TR>
		  <TD class=lightbluetable align=right>所在地區(Area)</TD>
		  <TD class=whitetablebg>
			<Select name="tfx_mref" size=1>
			<option value="">--- 請選擇 ---</option>
			<%
				sqlCom ="SELECT * FROM CodeMain Where codeMetaId='area' ORDER BY msortValue"
					'response.write SqlCom
				set RSS = conn.execute(sqlCom)
						While not RSS.EOF%>
						<option value="<%=RSS("mcode")%>"><%=RSS("mvalue")%></option>
						<%	RSS.movenext
						wend%>
			%>
			</select>
		  </TD>
		</TR>
		<%end if%>
</TABLE></CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
<% if formFunction = "add" then %>
      <% if (HTProgRight and 4)=4 then %><input type="button" value="新增存檔" name="task" class="cbutton" OnClick="VBS:formAdd">
      <input type=button value="清除重填" class="cbutton" OnClick="VBS:formReset"><% End If %>         
<% elseif formFunction = "edit" then %>
      <% if (HTProgRight and 8)=8 then %><input type="button" value="編修存檔" name="task" class="cbutton" OnClick="VBS:formEdit">
      <input type=button value="清除重填" class="cbutton" OnClick="VBS:formReset"><% End If %>&nbsp;       
      <% if (HTProgRight and 16)=16 then %><input type="button" value="刪除" name="task" class="cbutton" OnClick="VBS:formDelete"><%end if%>
<% End If %>
 </td></tr>
</table>
