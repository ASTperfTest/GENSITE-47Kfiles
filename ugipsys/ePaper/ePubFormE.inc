<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="ePubEdit.asp?epubId=<%=request.querystring("epubId")  %>">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="eTableLable" align="right"><font color="red">*</font>
發行日期：</TD>
<TD class="eTableContent"><input name="htx_pubDate" type=hidden>
<input name="pcShowhtx_pubDate" size="8" readonly onclick="VBS: popCalendar 'htx_pubDate' ,''">
<input type="hidden" name="htx_ePubID">
<input type="hidden" name="epTreeID" value="<%=Request("epTreeID")%>">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
標題：</TD>
<TD class="eTableContent"><input name="htx_title" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
資料範圍起日：</TD>
<TD class="eTableContent"><input name="htx_dbDate" type=hidden>
<input name="pcShowhtx_dbDate" size="8" readonly onclick="VBS: popCalendar 'htx_dbDate' ,''">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
資料範圍迄日：</TD>
<TD class="eTableContent"><input name="htx_deDate" type=hidden>
<input name="pcShowhtx_deDate" size="8" readonly onclick="VBS: popCalendar 'htx_deDate' ,''">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
資料則數：</TD>
<TD class="eTableContent"><input name="htx_maxNo" size="4" title="各單元內最多幾筆資料">
</TD>
</TR>
<TR>
	<TD class="eTableLable" align="right"><font color="red">*</font>
	版型選擇：</TD>
	<TD class="eTableContent"><input type="radio" id="type1" name="htx_pubType"  value="0" /><label for="type1">版型1</label>
							  <input type="radio" id="type2" name="htx_pubType" value="1" /><label for="type2">版型2</label>
							  <input type="radio" id="type3" name="htx_pubType" value="2" /><label for="type3">版型3</label>
	</TD>
</TR>
<TR>
<TD class="eTableLable" align="right">
	試發用mail：</TD>
	<TD class="eTableContent">
	<input name="htx_ePubSendDemoMail" size="30" title="試發用mail">
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
            <input type=button value ="查　詢" class="cbutton" onClick="formSearchSubmit()">
            <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <%end if%>    
<%Elseif formFunction = "edit" then %>
        <%if (HTProgRight AND 8) <> 0 then %>
               <input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()">
        <% End If %>
        
       <% if (HTProgRight AND 16) <> 0 then %>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>           
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">

<% Else '-- add ---%>          
          <%if (HTProgRight AND 4)<>0 then %>
               <input type=button value ="新增存檔" class="cbutton" onClick="formAddSubmit()">
          <%end if%>     
          
          <%if (HTProgRight AND 4)<>0  then %>
              <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
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
Dim followCanTarget

sub popCalendar(dateName,followName)        
 	CanTarget=dateName
 	followCanTarget=followName
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
        document.all("pcShow"&CanTarget).value=document.all.CalendarTarget.value
        if followCanTarget<>"" then
        	document.all(followCanTarget).value=document.all.CalendarTarget.value
        	document.all("pcShow"&followCanTarget).value=document.all.CalendarTarget.value
        end if
	end if
end sub   
</script>
