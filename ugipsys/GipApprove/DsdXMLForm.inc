<form method="POST" name="reg" action="" ENCTYPE="MULTIPART/FORM-DATA">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER>
<TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<%	
	for each param in allModel2.selectNodes("//fieldList/field[formList!='']") 
	      	if nullText(param.selectSingleNode("inputType"))<>"hidden" and nullText(param.selectSingleNode("fieldName"))<>"iDept" then
		    response.write "<TR><TD class=""eTableLable"" align=""right"">"
		    if nullText(param.selectSingleNode("canNull")) = "N" then _
			response.write "<font color=""red"">*</font>"
		    response.write nullText(param.selectSingleNode("fieldLabel")) & "</TD>"
		    response.write "<TD class=""eTableContent"">"
	      	end if
		if formFunction = "edit" AND nullText(param.selectSingleNode("isPrimaryKey"))="Y" then 
			orgInputType = param.selectSingleNode("inputType").text
			param.selectSingleNode("inputType").text = "readonly"
			processParamField param
			param.selectSingleNode("inputType").text = orgInputType
		elseif nullText(param.selectSingleNode("fieldName"))="fCTUPublic" then
			param.selectSingleNode("refLookup").text = "isPublic3"
			processParamField param		
			param.selectSingleNode("refLookup").text = "isPublic"				
		elseif nullText(param.selectSingleNode("fieldName"))="iDept" then
		    if (HTProgRight AND 64) <> 0 then	
		        response.write "<TR><TD class=""eTableLable"" align=""right"">"
		        if nullText(param.selectSingleNode("canNull")) = "N" then _
				response.write "<span class=""Must"">*</span>"
		        response.write nullText(param.selectSingleNode("fieldLabel")) & "</TD>"
		        response.write "<TD class=""eTableContent"">"		    	
			response.write "<Select name=""htx_iDept"" size=1>" & vbCRLF	
			SQL="Select deptID,deptName from Dept WHERE inUse='Y' Order by kind, deptID"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>
			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend
			response.write "</select>"				    
			response.write "</TD></TR>"				    
		    else
			response.write "<input type=""hidden"" name=""htx_iDept"">" & vbCRLF
		    end if		    	
		else
		    	processParamField param		    
		end if
		if session("CodeID")<>"7" AND nullText(param.selectSingleNode("fieldName")) = "xImportant" then
%>
		&nbsp;　&nbsp;<INPUT type="checkbox" name="xxCheckImportant" value="Y">顯示於每日頭條
<%
		end if
	      	if nullText(param.selectSingleNode("fieldName"))="xKeyword" then	
		    response.write "<input type=""button"" value=""自動產生"" class=""cbutton"" onClick=""VBS:keywordMake"">"
	     	end if		
	      	if nullText(param.selectSingleNode("inputType"))<>"hidden" and nullText(param.selectSingleNode("fieldName"))<>"iDept" then
		    response.write "</TD></TR>"
	     	end if
	next
%>
</TABLE>
</CENTER>
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
   		<OPTION value="CuCheckList.asp">列印待審稿</OPTION>   		
   		</SELECT> &nbsp;
          <%if (HTProgRight AND 4)<>0 then %>
               <input type=button value ="新增存檔" class="cbutton" onClick="formAddSubmit()">
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
