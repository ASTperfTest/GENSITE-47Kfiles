<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="lightbluetable" align="right">�����Ҹ��G</TD>
<TD class="whitetablebg"><input name="htx_psnID" size="10" readonly="true" class="rdonly">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">�m�W�G</TD>
<TD class="whitetablebg"><input name="htx_pName" size="20" readonly="true" class="rdonly">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">�X�ͤ�G</TD>
<TD class="whitetablebg"><input name="htx_birthDay" size="10" readonly onclick="VBS: popCalendar 'htx_birthDay'">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">�A�ȳ��G</TD>
<TD class="whitetablebg"><input name="htx_myOrg" size="50" readonly="true" class="rdonly">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">���}�G</TD>
<TD class="whitetablebg"><input name="htx_addr" size="50" readonly="true" class="rdonly">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">eMail�G</TD>
<TD class="whitetablebg"><input name="htx_eMail" size="50" readonly="true" class="rdonly">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">�s���q�ܡG</TD>
<TD class="whitetablebg"><input name="htx_tel" size="20" readonly="true" class="rdonly">
</TD>
</TR>

<TR><TD class="whitetablebg" colspan="2"><HR/></TD>
</TR>
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>
���A�G</TD>
<TD class="whitetablebg"><Select name="htx_status" size=1>
			<%SQL="Select mCode,mValue from CodeMain where mSortValue IS NOT NULL AND codeMetaID='ppCEstatus' Order by mSortValue"
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

                <input type=button value ="�s�צs��" class="cbutton" onClick="formModSubmit()">
       
             <input type=button value ="�R�@��" class="cbutton" onClick="formDelSubmit()">   

        <input type=button value ="���@��" class="cbutton" onClick="resetForm()">
        <input type=button value ="�^�e��" class="cbutton" onClick="Vbs:history.back">


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