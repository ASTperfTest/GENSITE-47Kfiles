<form method="POST" name="reg" action="">
<%if formFunction<>"edit" then%>
<p align="Left"><font size="2">
多媒體物件存放站台目錄新增說明:<br>
  1.目錄名稱(實體路徑用)欄位,請輸入英文字母字元或數字,內無空白字元或其他特殊符號<br>
</font></p>
<%end if%>
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=pfx_MMOFolderID value="">
<INPUT TYPE=hidden name=xxPath value="<%=xxPath%>">
<INPUT TYPE=hidden name=MMOSiteID value="<%=xMMOSiteID%>">
<CENTER><TABLE border=0 class=bluetable cellspacing=1 cellpadding=2 width="70%">
		<TR>
		  <TD class=lightbluetable align=right>目錄名稱(實體路徑用)<font color=red>*</font></TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_MMOFolderName SIZE=15 <%if formFunction="edit" then%> readonly="true" class="rdonly"<%end if%>></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>目錄顯示名稱<font color=red>*</font></TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_MMOFolderNameShow SIZE=15 ></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>目錄說明</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_MMOFolderDesc SIZE=30></TD>
		</TR>
		<%if (HTProgRight AND 64) <> 0 and deptCheck then%>
		<TR><TD class="lightbluetable" align="right">單 位：</TD>
		<TD class="whitetablebg">
		<Select name="htx_deptID" size=1>
		<option value="">--- 共 用 ---</option>
		<%
			sqlCom ="SELECT D.deptID,D.deptName,D.parent,len(D.deptID)-1,D.seq," & _
				"(Select Count(*) from Dept WHERE parent=D.deptID AND nodeKind='D') " & _
				"  FROM Dept AS D Where D.nodeKind='D' " _
				& " AND D.deptID LIKE '" & session("deptID") & "%'" _
				& " ORDER BY len(D.deptID), D.parent, D.seq"	
		'response.write SqlCom
			set RSS = conn.execute(sqlCom)
			if not RSS.EOF then
				ARYDept = RSS.getrows(200)
				glastmsglevel = 0
				genlist 0, 0, 1, 0
			        expandfrom ARYDept(cid, 0), 0, 0
			end if
		%>
		</select>
		</TD></TR>
		<%elseif formFunction ="edit" and not deptCheck then%>
			<TR><TD class="lightbluetable" align="right">單 位：</TD>
			<TD class="whitetablebg"><%=RSreg("deptName")%></TD></TR>
			<input type="hidden" name="htx_deptID" value="">
		<%else%>
			<input type="hidden" name="htx_deptID" value="">
		<%end if%>
</TABLE></CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
<% if formFunction = "add" then %>
      <% if (HTProgRight and 4)=4 then %>
      	<input type="button" value="新增存檔" name="task" class="cbutton" OnClick="VBS:formAdd">
      	<input type=button value="清除重填" class="cbutton" OnClick="VBS:formReset">
      <% End If %>         
<% End If %>
<% if formFunction = "edit" and deptCheck then %>
      <% if (HTProgRight and 8)=8 then %>
      	<input type="button" value="編修存檔" name="task" class="cbutton" OnClick="VBS:formEdit">
      	<input type=button value="重  填" class="cbutton" OnClick="VBS:formReset">
      <% End If %>         
       <% if (HTProgRight AND 16) <> 0 and RSreg("MMOFolderChildCount") = 0 and MMODataCount = 0 then %>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>  
<% End If %>
 </td></tr>
</table>
