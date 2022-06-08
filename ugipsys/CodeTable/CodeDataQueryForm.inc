<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border=0 class=bluetable cellspacing=1 cellpadding=2 width="70%">
		<TR>
		  <TD class=lightbluetable align=right><%=RSCode("CodeValueFldName")%>(Value)</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_CodeValue SIZE=15></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right><%=RSCode("CodeDisplayFldName")%>(Display)</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_CodeDisplay SIZE=30></TD>
		</TR>
		<%if session("SortFlag") then%>
		<TR>
		  <TD class=lightbluetable align=right><%=RSCode("CodeSortFldName")%>(SORT)</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_CodeSort SIZE=15></TD>
		</TR>
		<%end if%>		
</TABLE></CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
<% if formFunction = "query" then %>
      <% if (HTProgRight and 2)=2 then %><input type="button" value="確定查詢" name="task" class="cbutton" OnClick="VBS:formSearch">
      <input type=button value="清除重填" class="cbutton" OnClick="VBS:resetform"><% End If %>         
<% End If %>
 </td></tr>
</table>
