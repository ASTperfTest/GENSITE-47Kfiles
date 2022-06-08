<form method="POST" name="reg"  ENCTYPE="MULTIPART/FORM-DATA" action="CuMMOAdd.asp">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER>
<table width="100%" cellspacing="0" id="PopList">
<%	
	for each param in allModel2.selectNodes("//fieldList/field[formList!='']") 
	      	if nullText(param.selectSingleNode("inputType"))<>"hidden" and nullText(param.selectSingleNode("fieldName"))<>"idept" and nullText(param.selectSingleNode("fieldName"))<>"mmofolderId" then
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
		elseif nullText(param.selectSingleNode("fieldName"))="idept" then
		    if (HTProgRight AND 64) <> 0 then	
		        response.write "<TR><TD class=""Label"" align=""right"">"
		        if nullText(param.selectSingleNode("canNull")) = "N" then _
				response.write "<span class=""Must"">*</span>"
		        response.write nullText(param.selectSingleNode("fieldLabel")) & "</TD>"
		        response.write "<TD class=""eTableContent"">"		
				response.write "<Select name=""htx_iDept"" size=1>" & vbCRLF	
			sqlCom ="SELECT D.deptId,D.deptName,D.parent,len(D.deptId)-1,D.seq," & _
				"(Select Count(*) from Dept WHERE parent=D.deptId AND nodeKind='D') " & _
				"  FROM Dept AS D Where D.nodeKind='D' " _
				& " AND D.deptId LIKE '" & session("deptID") & "%'" _
				& " ORDER BY len(D.deptId), D.parent, D.seq"	
			'----*)
			set RSS = conn.execute(sqlCom)
			if not RSS.EOF then
				ARYDept = RSS.getrows(200)
				glastmsglevel = 0
				genlist 0, 0, 1, 0
			        expandfrom ARYDept(cid, 0), 0, 0
			end if
			response.write "</select>"				    
		    else
			response.write "<input type=""hidden"" name=""htx_iDept"">" & vbCRLF
		    end if		    	
		elseif nullText(param.selectSingleNode("fieldName"))="mmofolderId" then
			SQLM="Select MM.mmositeId+MM.mmofolderName as MMOFOlderID,MM.mmositeId,MM.mmofolderName,MS.upLoadSiteFtpip,MS.upLoadSiteFtpport,MS.upLoadSiteFtpid,MS.upLoadSiteFtppwd " & _
				"from CtUnit C Left Join Mmofolder MM ON C.mmofolderId=MM.mmofolderId " & _
				"	Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId " & _
				"where C.CtUnitID="&request("htx_CtUnitID")
			Set RSM=conn.execute(SQLM)
			if not RSM.eof then xMMOFolderID=RSM("MMOFOlderID")
		        response.write "<TR><TD class=""Label"" align=""right"">"
		        if nullText(param.selectSingleNode("canNull")) = "N" then _
				response.write "<span class=""Must"">*</span>"
		        response.write nullText(param.selectSingleNode("fieldLabel")) & "</TD>"
		        response.write "<TD class=""eTableContent"">"		    	
			response.write "<Select name=""htx_MMOFolderID"" size=1>" & vbCRLF	
		  	sqlCom="Select CASE mmofolderName WHEN 'zzz' THEN 0 ELSE mmofolderId END MMOFolderID, " & _
		  		"Case MM.mmofolderParent when 'zzz' then MS.mmositeName else MM.mmofolderNameShow END MMOFolderNameShow, " & _
		  		"Case mmofolderParent when 'zzz' then 0 else (Select mmofolderId from Mmofolder where mmositeId=MM.mmositeId and mmofolderName=MM.mmofolderParent) END MMOFolderParent " & _
		  		",1 " & _
				",1 " & _
				",Case mmofolderName when 'zzz' then " & _
				"	(Select count(*) from Mmofolder MM2 Left Join Mmosite MS2 ON MM2.mmositeId=MS2.mmositeId where (MS2.deptId IS NULL OR MS2.deptId LIKE '" & session("deptID") & "%' OR MS2.deptId = Left('" & session("deptID") & "',Len(MS2.deptId))) and (MM2.deptId IS NULL OR MM2.deptId = Left('" & session("deptID") & "',Len(MM2.deptId)) OR MM2.deptId like '"&session("deptID")&"%') and MM2.mmofolderParent='zzz') " & _
				" else " & _
				"	(Select count(*) from Mmofolder MM2 Left Join Mmosite MS2 ON MM2.mmositeId=MS2.mmositeId where (MS2.deptId IS NULL OR MS2.deptId LIKE '" & session("deptID") & "%' OR MS2.deptId = Left('" & session("deptID") & "',Len(MS2.deptId))) and (MM2.deptId IS NULL OR MM2.deptId = Left('" & session("deptID") & "',Len(MM2.deptId)) OR MM2.deptId like '"&session("deptID")&"%') and MM2.mmofolderParent=MM.mmofolderName and MM2.mmositeId=MM.mmositeId) " & _
				" END childCount " & _
				"from Mmofolder MM Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId " & _
				"where (MS.deptId IS NULL OR MS.deptId LIKE '" & session("deptID") & "%' OR MS.deptId = Left('" & session("deptID") & "',Len(MS.deptId))) and (MM.deptId IS NULL OR  MM.deptId = LEFT('" & session("deptID") & "', Len(MM.deptId)) OR MM.deptId like '"&session("deptID")&"%') "		
			if xMMOFolderID<>"" then 
				sqlCom=sqlCom & " and CASE mmofolderName WHEN 'zzz' THEN mmofolderName ELSE MM.mmositeId + mmofolderName END like '"&xMMOFolderID&"%' "
			end if
			sqlCom=sqlCom & " order by Case mmofolderParent when 'zzz' then '' else mmofolderParent END, mmofolderID"
'				response.write sqlCom
			set RSS = conn.execute(sqlCom)
			if not RSS.EOF then
				ARYDept = RSS.getrows(200)
				glastmsglevel = 0
				genlist 0, 0, 1, 0
			        expandfrom ARYDept(cid, 0), 0, 0
			end if
			response.write "</select>"
			response.write "<input type=button value =""新增子目錄"" class=""cbutton"" onClick=""MMOFolderAdd()"">"
			response.write "</TD></TR>"	
		else
		    	processParamField param		    
		end if
		if session("CodeID")<>"7" AND nullText(param.selectSingleNode("fieldName")) = "xImportant" then
%>
		&nbsp;　&nbsp;<INPUT type="checkbox" name="xxCheckImportant" value="Y">顯示於每日頭條
<%
		end if
	      	if nullText(param.selectSingleNode("inputType"))<>"hidden" then	
		    response.write "</TD></TR>"
	     	end if
	next
%>
</TABLE>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
          <%if (HTProgRight AND 4)<>0 then %>
               <input type=button value ="新增存檔" class="InputButton" onClick="formAddSubmit()">
              <input type=button value ="重　填" class="InputButton" onClick="resetForm()">
          <%end if%>    
	<input type="button" class="InputButton" value="回查詢" onClick="Javascript:window.navigate('CuMMOQuery.asp');">

 </td></tr>
</table>
</CENTER>
<object data="../inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
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
