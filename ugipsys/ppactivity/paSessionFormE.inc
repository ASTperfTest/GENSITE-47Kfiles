<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="lightbluetable" align="right">�ҵ{�s���G</TD>
<TD class="whitetablebg"><input name="htx_paSnum" size="30">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">�ҵ{�ɶ��G</TD>
<TD class="whitetablebg"><input name="htx_dtNote" size="50">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">�ҵ{�a�I�G</TD>
<TD class="whitetablebg"><input name="htx_place" size="50">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">���W�I���G</TD>
<TD class="whitetablebg"><input name="htx_bDate" size="10" readonly onclick="VBS: popCalendar 'htx_bDate'">
<input type="hidden" name="htx_actID">
<input type="hidden" name="htx_paSID">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">�H�ƭ���G</TD>
<TD class="whitetablebg"><input name="htx_pLimit" size="2">
</TD>
</TR>

<TR><TD class="lightbluetable" align="right">�ҵ{�p���H��T�G</TD>
<TD class="whitetablebg"><input name="htx_refPage" size="30">
</TD>
</TR>
</TR>
<TR><TD class="lightbluetable" align="right">�ҵ{���A�G</TD>
<TD class="whitetablebg"><Select name="htx_aStatus" size=1>
<option value="">�п��</option>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID='classStatus' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
</TABLE>
</CENTER>
<object data="../inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
<%If formFunction = "query" then %>
        <%if (HTProgRight AND 2) <> 0 then %>
            <input type=button value ="�d�@��" class="cbutton" onClick="formSearchSubmit()">
            <input type=button value ="���@��" class="cbutton" onClick="resetForm()">
        <%end if%>    
<%Elseif formFunction = "edit" then %>
        <%if (HTProgRight AND 8) <> 0 then %>
               <input type=button value ="�s�צs��" class="cbutton" onClick="formModSubmit()">
        <% End If %>
        
       <% if (HTProgRight AND 16) <> 0 then %>
             <input type=button value ="�R�@��" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>           
        <input type=button value ="���@��" class="cbutton" onClick="resetForm()">
        <input type=button value ="�^�e��" class="cbutton" onClick="Vbs:history.back">

<% Else '-- add ---%>          
          <%if (HTProgRight AND 4)<>0 then %>
               <input type=button value ="�s�W�s��" class="cbutton" onClick="formAddSubmit()">
          <%end if%>     
          
          <%if (HTProgRight AND 4)<>0  then %>
              <input type=button value ="���@��" class="cbutton" onClick="resetForm()">
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
		<script language=vbs>
Dim CanTarget

sub popCalendar(dateName)        
 	CanTarget=dateName
	xdate = document.all(CanTarget).value
	if not isDate(xdate) then	xdate = date()
	document.all.calendar1.setDate xdate
	
 	If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 	Else        
       ex=window.event.clientX
       ey=document.body.scrolltop+window.event.clientY+10
       if ex>520 then ex=520
       document.all.calendar1.style.pixelleft=ex-80
       document.all.calendar1.style.pixeltop=ey
       document.all.calendar1.style.visibility=""
 	End If              
end sub     

sub calendar1_onscriptletevent(n,o) 
    	document.all("CalendarTarget").value=n         
    	document.all.calendar1.style.visibility="hidden"    
    if n <> "Cancle" then
        document.all(CanTarget).value=document.all.CalendarTarget.value
	end if
end sub   
</script>