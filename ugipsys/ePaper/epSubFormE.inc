<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
訂閱 Email：</TD>
<TD class="eTableContent"><input name="htx_eMail" type=hidden>
<%=RSreg("email")%>
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
訂閱日：</TD>
<TD class="eTableContent"><%=RSreg("createtime")%>
</TD>
</TR>
</TABLE>
</CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
<%if formFunction = "edit" then %>
        
       <% if (HTProgRight AND 16) <> 0 then %>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>           
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">

<% End If %>
 </td></tr>
</table>
