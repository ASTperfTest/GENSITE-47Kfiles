<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="eTableLable" align="right"><font color="red">*</font>
登入密碼：</TD>
<TD class="eTableContent"><input name="htx_myPassword" type="password" size="12">
 密碼確認：
<input name="htx_myPassword2check" type="password" size="12" title="再輸入一次密碼供確認檢查">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
身份證號：</TD>
<TD class="eTableContent"><input name="htx_psnID" size="10">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
姓名：</TD>
<TD class="eTableContent"><input name="htx_pName" size="20">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
出生日：</TD>
<TD class="eTableContent"><input name="htx_birthDay" type=hidden>
<input name="pcShowhtx_birthDay" size="10" readonly onclick="VBS: popCalendar 'htx_birthDay' ,''">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">eMail：</TD>
<TD class="eTableContent"><input name="htx_eMail" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
連絡電話：</TD>
<TD class="eTableContent"><input name="htx_tel" size="20">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
聯絡地址：</TD>
<TD class="eTableContent"><input name="htx_emergContact" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">工作部門：</TD>
<TD class="eTableContent"><input name="htx_deptName" size="20">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">職稱：</TD>
<TD class="eTableContent"><input name="htx_jobName" size="20">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">最高學歷：</TD>
<TD class="eTableContent"><Select name="htx_topEdu" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from codeMain where mSortValue IS NOT NULL AND codeMetaID='eduCode' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right">葷/素：</TD>
<TD class="eTableContent"><Select name="htx_meatKind" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from codeMain where mSortValue IS NOT NULL AND codeMetaID='meatKind' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right">公司全銜：</TD>
<TD class="eTableContent"><input name="htx_corpName" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">統一編號：</TD>
<TD class="eTableContent"><input name="htx_corpID" size="8">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">聯絡地址：</TD>
<TD class="eTableContent"><input name="htx_corpAddr" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">聯絡人：</TD>
<TD class="eTableContent"><input name="htx_corpContact" size="20">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">電話：</TD>
<TD class="eTableContent"><input name="htx_corpTel" size="20">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">傳真：</TD>
<TD class="eTableContent"><input name="htx_corpFax" size="20">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">消息來源：</TD>
<TD class="eTableContent"><input name="htx_datafrom" size="20">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">行動電話：</TD>
<TD class="eTableContent"><input name="htx_mobile" size="20">
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
