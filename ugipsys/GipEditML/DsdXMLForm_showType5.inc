<form id="Form1" method="POST" name="reg" action="" ENCTYPE="MULTIPART/FORM-DATA">
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=showType value="">
<INPUT TYPE=hidden name=refID value="">
<TABLE cellspacing="0">
<%	
	for each param in allModel2.selectNodes("//fieldList/field[formList!='']") 
	    if not (nullText(param.selectSingleNode("fieldName"))="fCTUPublic" and session("CheckYN")="Y") then
	      	if nullText(param.selectSingleNode("inputType"))<>"hidden" then
		    response.write "<TR><TD class=""Label"" align=""right"">"
		    if nullText(param.selectSingleNode("canNull")) = "N" then _
			response.write "<span class=""Must"">*</span>"
		    response.write nullText(param.selectSingleNode("fieldLabel")) & "</TD>"
		    if nullText(param.selectSingleNode("fieldRefEditYN"))="Y" then
		    	response.write "<TD class=""eTableContent"">"
		    	processParamField param	
		    	response.write "</TD></TR>"		    		    		    	
		    else
		    	response.write "<TD class=""eTableContent"" id=""x"&nullText(param.selectSingleNode("fieldName"))&"""></TD></TR>"	
			orgInputType = param.selectSingleNode("inputType").text
			param.selectSingleNode("inputType").text = "hidden"
			processParamField param
			param.selectSingleNode("inputType").text = orgInputType		    		    
		    end if
		else
		    	processParamField param		    
	      	end if
	    end if
	next
%>
</TABLE>
<object data="../inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
   		<font size=2>接續處理：</font><SELECT NAME=nextTask size=1>
   		<OPTION value="" selected></OPTION>
   		<OPTION value="CuAttachList.asp">附件</OPTION>
		<% if session("CodeID")="7" then %>
   		<OPTION value="CuPageList.asp">連結</OPTION>
   		<% end if %>
   		<OPTION value="CuHtmlEdit.asp">網頁</OPTION>
		<% if session("CheckYN")="Y" then %>   		
   		<OPTION value="CuCheckList.asp">列印待審稿</OPTION>   	
   		<% end if %>	
   		<OPTION value="DsdXMLAdd_Multi.asp">多向出版</OPTION>
      		</SELECT> &nbsp;
          <%if (HTProgRight AND 4)<>0 then %>
              <%if nullText(allModel.selectSingleNode("//sBaseDSDXML/addYN"))="" then %>          
		<input type=button value ="新增存檔" class="cbutton" onClick="formAddSubmit()">
	      <%end if%> 
          <%end if%>     
          　
          <%if (HTProgRight AND 4)<>0  then %>
              <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
          <%end if%>    

 </td></tr>
</table>
</form>     
<SCRIPT LANGUAGE="vbs">
    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub

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
