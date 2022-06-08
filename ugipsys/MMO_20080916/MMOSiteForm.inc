<form method="POST" name="reg" action="">
<p align="Left"><font size="2">
<%if formFunction="edit" then%>
編修多媒體物件存放站台說明:<br>
  1.站台下若存在子目錄時,不允許刪除<br>
  2.測試FTP同步站台是否正常連接,請按[ 測試FTP ]按鈕
<%else%>
新增多媒體物件存放站台說明:<br>
  1.本機GIP後台管理網站/site/xxx/public下(xxx為session("mySiteID")值),將以[ 站台ID ]欄位<br>　內容值為目錄名稱,新增一個子目錄<br>
  2.若要建立多媒體物件的FTP同步存放站台,請架好FTP同步存放站台(寫入權限打開),所指根目錄下先建立<br>　public目錄夾
    (請確定該FTP站台已經可以運作),再填入以下[ FTP同步站台 IP/Port/登入帳號/登入密碼 ]<br>　等四個欄位, 
    存檔後除新增本機目錄外,將會在該FTP站台/public下,以[ 站台ID ]欄位內容值為目錄名稱,<br>　新增一個子目錄<br>
<%end if%>
</font></p>
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border=0 class=bluetable cellspacing=1 cellpadding=2 width="70%">
		<TR>
		  <TD class=lightbluetable align=right>站台ID<font color=red>*</font></TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=htx_MMOSiteID SIZE=15 <%if formFunction="edit" then%> readonly="true" class="rdonly"<%end if%>></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>站台名稱<font color=red>*</font></TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=htx_MMOSiteName SIZE=30></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>站台說明</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=htx_MMOSiteDesc SIZE=30></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>FTP同步站台 IP</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=htx_UpLoadSiteFTPIP SIZE=30></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>FTP同步站台 Prot</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=htx_UpLoadSiteFTPPort SIZE=30></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>FTP同步站台登入帳號</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=htx_UpLoadSiteFTPID SIZE=30></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>FTP同步站台登入密碼</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=htx_UpLoadSiteFTPPWD SIZE=30></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>多媒體物件連結Domain</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=htx_UpLoadSiteHTTP SIZE=30></TD>
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
				ARYDept = RSS.getrows(300)
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
<% if formFunction = "edit" then %>
	<%if deptCheck then%>
      <% if (HTProgRight and 8)=8 then %>
      	<input type="button" value="編修存檔" name="task" class="cbutton" OnClick="VBS:formEdit">
      	<input type=button value="重  填" class="cbutton" OnClick="VBS:formReset">
      <% End If %>  
       <% if (HTProgRight AND 16) <> 0 and RSreg("MMOSiteChildCount") <= 1 then %>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>  
       <%end if%>
      <input type=button name=FTPTest value="測試FTP" class="cbutton" OnClick="VBS:FTPTest">       
<% End If %>
 </td></tr>
</table>
