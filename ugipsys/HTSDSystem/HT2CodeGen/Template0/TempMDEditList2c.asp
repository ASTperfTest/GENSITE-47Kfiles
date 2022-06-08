<BR>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
        <%if (HTProgRight AND 8) <> 0 then %>
               <input type=button value ="½s­×¦sÀÉ" class="cbutton" onClick="formModSubmit()">
        <% End If %>
        
       <% if (HTProgRight AND 16) <> 0 and deleteFlag then %>
             <input type=button value ="§R¡@°£" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>           
        <input type=button value ="­«¡@¶ñ" class="cbutton" onClick="resetForm()">
 </td></tr>
</table>