<form id="Form1" method="POST" name="reg" action="DsdXMLAdd.asp" ENCTYPE="MULTIPART/FORM-DATA">
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=showType value="">
<INPUT TYPE=hidden name=refID value="">
<TABLE cellspacing="0">
<%	
	for each param in allModel2.selectNodes("//fieldList/field[formList!='']") 
	    if not (nullText(param.selectSingleNode("fieldName"))="fctupublic" and session("CheckYN")="Y") then
	      	if nullText(param.selectSingleNode("inputType"))<>"hidden" and nullText(param.selectSingleNode("fieldName"))<>"idept" and nullText(param.selectSingleNode("fieldName"))<>"mmofolderId" then
		    response.write "<TR><TD class=""Label"" align=""right"">"
		    if nullText(param.selectSingleNode("canNull")) = "N" then _
			response.write "<span class=""Must"">*</span>"
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
			response.write "<Select name=""htx_idept"" size=1>" & vbCRLF	
			sqlCom ="SELECT D.deptId,D.deptName,D.parent,len(D.deptId)-1,D.seq," & _
				"(Select Count(*) from Dept WHERE parent=D.deptId AND nodeKind='D') " & _
				"  FROM Dept AS D Where D.nodeKind='D' " _
				& " AND D.deptId LIKE '" & session("deptId") & "%'" _
				& " ORDER BY len(D.deptId), D.parent, D.seq"	
			'response.write SqlCom
			set RSS = conn.execute(sqlCom)
			if not RSS.EOF then
				ARYDept = RSS.getrows(300)
				glastmsglevel = 0
				genlist 0, 0, 1, 0
			        expandfrom ARYDept(cid, 0), 0, 0
			end if
			response.write "</select>"
			response.write "</TD></TR>"				    
		    else
			response.write "<input type=""hidden"" name=""htx_iDept"">" & vbCRLF
		    end if
		elseif nullText(param.selectSingleNode("fieldName"))="mmofolderId" then
			SQLM="Select MM.mmositeId+MM.mmofolderName as MMOFOlderID,MM.mmositeId,MM.mmofolderName,MS.upLoadSiteFtpip,MS.upLoadSiteFtpport,MS.upLoadSiteFtpid,MS.upLoadSiteFtppwd " & _
				"from CtUnit C Left Join Mmofolder MM ON C.mmofolderId=MM.mmofolderId " & _
				"	Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId " & _
				"where C.CtUnitId='"&session("CtUnitID")&"'"
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
				ARYDept = RSS.getrows(300)
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
'------- checkGIPconfig("xxCheckImportant") -------- start --------------------------------------------------------	
		if checkGIPconfig("KeywordAutoGen") then
			if session("CodeID")<>"7" AND nullText(param.selectSingleNode("fieldName")) = "xImportant" then
	%>
			&nbsp;　&nbsp;<INPUT type="checkbox" name="xxCheckImportant" value="Y">顯示於每日頭條
	<%
			end if
		end if	
'------- checkGIPconfig("xxCheckImportant") -------- end --------------------------------------------------------	
'------- checkGIPconfig("KeywordAutoGen") -------- start --------------------------------------------------------	
		if checkGIPconfig("KeywordAutoGen") then
		      	if nullText(param.selectSingleNode("fieldName"))="xkeyword" then	
			    response.write "<input type=""button"" value=""自動產生"" class=""cbutton"" onClick=""VBS:keywordMake"">"
		     	end if		
		end if	
'------- checkGIPconfig("KeywordAutoGen") -------- end --------------------------------------------------------	
	      	if nullText(param.selectSingleNode("inputType"))<>"hidden" and nullText(param.selectSingleNode("fieldName"))<>"iDept" then
		    response.write "</TD></TR>"
	     	end if
	    end if
	next
%>
</TABLE>
<object data="../inc/calendar.asp" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
   		<font size=2>接續處理：</font><SELECT NAME=nextTask size=1>
   		<OPTION value="" selected></OPTION>
   		<OPTION value="CuAttachList.asp?iCuItem=">附件</OPTION>
		<% if session("CodeID")="7" then %>
   		<OPTION value="CuPageList.asp?iCuItem=">連結</OPTION>
   		<% end if %>
   		<OPTION value="CuHtmlEdit.asp?iCuItem=">網頁</OPTION>
		<% if session("CheckYN")="Y" then %>   		
   		<OPTION value="CuCheckList.asp?iCuItem=">列印待審稿</OPTION>   	
   		<% end if %>	
   		
   	<% 	set xhtPageDom = session("codeXMLSpec")
   		for each xop in xhtPageDom.selectNodes("//addOpURL") %>
   		<OPTION value="<%=nullText(xop.selectSingleNode("opURL"))%>"><%=nullText(xop.selectSingleNode("opName"))%></OPTION>
   	<% next	%>
   		<!--OPTION value="DsdXMLAdd_Multi.asp">多向出版</OPTION-->
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
