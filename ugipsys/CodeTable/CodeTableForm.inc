<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border=0 class=bluetable cellspacing=1 cellpadding=2 width="450">
		<TR>
		  <TD class=lightbluetable align=right><font color=red>※</font>代碼ID</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=pfx_codeId SIZE=15<%IF request.querystring("codeId") <> "" then %> class=sedit readonly<% End If %>></TD>
		  <TD class=lightbluetable align=right><font color=red>※</font>代碼名稱</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_codeName SIZE=15></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right><font color=red>※</font>存放資料表</TD>
		  <TD class=whitetablebg colspan=3><Input NAME=tfx_codeTblName SIZE=30>
		  </TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>參照值ITEM欄</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_codeSrcFld SIZE=15></TD>
		  <TD class=lightbluetable align=right>參照值ITEM碼</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_codeSrcItem SIZE=15></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right><font color=red>※</font>Value值欄</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_codeValueFld SIZE=15></TD>
		  <TD class=lightbluetable align=right>Value值欄Heading</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_codeValueFldName SIZE=15></TD>		  
		</TR>
		<TR>
		  <TD class=lightbluetable align=right><font color=red>※</font>Display欄</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_codeDisplayFld SIZE=15></TD>
		  <TD class=lightbluetable align=right>Display欄Heading</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_codeDisplayFldName SIZE=15></TD>		  		  
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>過濾(排序)欄</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_codeSortFld SIZE=15></TD>
		  <TD class=lightbluetable align=right>過濾(排序)欄Heading</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_codeSortFldName SIZE=15></TD>		  		  
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>類別</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_codeType SIZE=15></TD>
		  <TD class=lightbluetable align=right>代碼ID順序</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_codeRank SIZE=15></TD>		  		  
		</TR>	
		<TR>
		  <TD class=lightbluetable align=right>是否show出?</TD>
		  <TD class=whitetablebg colspan=3><INPUT TYPE=checkbox NAME=bfx_ShowOrNot></TD>		  		  
		</TR>				
<%  if checkGIPconfig("codeML") then %>
		<TR>
		  <TD class=lightbluetable align=right>多國語言?</TD>
		  <TD class=whitetablebg colspan=3><INPUT TYPE=checkbox NAME=bfx_isML value="Y"></TD>		  		  
		</TR>				
<%	end if %>
		<TR>
		  <TD class=lightbluetable align=right>XML Spec</TD>
		  <TD class=whitetablebg colspan=3><Input NAME=tfx_codeXmlspec SIZE=40>
		  </TD>
		</TR>
</TABLE></CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
<% if formFunction = "edit" then %>
<% if (HTProgRight AND 8) <> 0 then %>
<input type=button value="編修存檔" class="cbutton" style="cursor:hand" onClick="formModSubmit()">
<% End If %>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=button value="清除重填" class="cbutton" style="cursor:hand" onClick="resetForm()">
<input type=button value="回上一頁" class="cbutton" onclick="vbs:history.back" style="cursor:hand">
<input type=button value="代碼資料維護" class="cbutton" onclick="vbs:window.navigate 'CodeEdit.asp?myid=<%=request.querystring("codeId")%>'" style="cursor:hand">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<% if (HTProgRight AND 16) <> 0 then %>
<input type=button value="刪除" class="cbutton" style="cursor:hand" onClick="formDelSubmit()">
<% End If %>
<% Else '-- add ---%>
<input type=button value="新增存檔" class="cbutton" style="cursor:hand" onClick="formAddSubmit()">
<input type=button value="清除重填" class="cbutton" style="cursor:hand" onClick="resetForm()">
<% End If %>
 </td></tr>
</table>
