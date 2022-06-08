<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border=0 class=bluetable cellspacing=1 cellpadding=2 width="400">
		<TR>
		  <TD class=lightbluetable align=right>代碼ID</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=pfx_CodeID SIZE=20 <%IF request.querystring("CodeID") <> "" then %> class=sedit readonly<% End If %>></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>代碼名稱</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_CodeName SIZE=20></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>存放資料表</TD>
		  <TD class=whitetablebg colspan=3><Input NAME=tfx_CodeTblName SIZE=20>
		  </TD>
		</TR>
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
