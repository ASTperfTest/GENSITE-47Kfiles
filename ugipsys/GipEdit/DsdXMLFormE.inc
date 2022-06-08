<%
' ============= Modified by Chris, 2006/08/24, to handle 由其它AP主控的資料(如單位)，上稿時應不能新增/刪除 ========================'
'		Document: 950822_智庫GIP擴充.doc
'  modified list:
'	『刪除』鈕 加 nullText(htPageDom.selectSingleNode("//sBaseDSDXML/addYN"))<>"N" 條件
%>
<form id="Form1" method="POST" name="reg" action="DsdXMLEdit.asp?icuitem=<%=request.querystring("icuitem")%>&keep=<%=request("keep")%>&nowpage=<%=request("nowpage")%>&pagesize=<%=request("pagesize")%>" ENCTYPE="MULTIPART/FORM-DATA">
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=showType value="">
<INPUT TYPE=hidden name=refID value="">
<table cellspacing="0">
<%
'if xShowType <> "5" then	

	for each param in allModel2.selectNodes("//fieldList/field[formList!='']") 
	    if not (nullText(param.selectSingleNode("fieldName"))="fctupublic" and session("CheckYN")="Y") then
	      	if nullText(param.selectSingleNode("inputType"))<>"hidden" and nullText(param.selectSingleNode("fieldName"))<>"idept" and nullText(param.selectSingleNode("fieldName"))<>"mmofolderId" then
		    response.write "<TR><TD class=""Label"" align=""right"">"
		    if nullText(param.selectSingleNode("canNull")) = "N" then _
			response.write "<span class=""Must"">*</span>"
		    response.write nullText(param.selectSingleNode("fieldLabel")) & "</TD>"
		    response.write "<TD class=""eTableContent"">"
	      	end if
	      	
		if formFunction = "edit" AND nullText(param.selectSingleNode("isPrimaryKey"))="Y" and nullText(param.selectSingleNode("identity"))<>"Y" then 
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
		        response.write "<TR><TD class=""Label"" align=""right"">"
		        if nullText(param.selectSingleNode("canNull")) = "N" then _
				response.write "<span class=""Must"">*</span>"
		        response.write nullText(param.selectSingleNode("fieldLabel")) & "</TD>"
		        response.write "<TD class=""eTableContent"">"		    	
			response.write "<input type=hidden name=""htx_MMOFolderID"">" & MMOPathStr(RSreg("mmofolderId")) & vbCRLF	
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
'else
'	for each param in allModel2.selectNodes("//fieldList/field[formList!='']") 
'	    if not (nullText(param.selectSingleNode("fieldName"))="fCTUPublic" and session("CheckYN")="Y") then
'	      	if nullText(param.selectSingleNode("inputType"))<>"hidden" then
'		    response.write "<TR><TD class=""Label"" align=""right"">"
'		    if nullText(param.selectSingleNode("canNull")) = "N" then _
'			response.write "<span class=""Must"">*</span>"
'		    response.write nullText(param.selectSingleNode("fieldLabel")) & "</TD>"
'		    if nullText(param.selectSingleNode("fieldRefEditYN"))="Y" then
'		    	response.write "<TD class=""eTableContent"">"
'		    	processParamField param	
'		    	response.write "</TD></TR>"		    		    		    	
'		    else
'		    	response.write "<TD class=""eTableContent"" id=""x"&nullText(param.selectSingleNode("fieldName"))&"""></TD></TR>"	
'			orgInputType = param.selectSingleNode("inputType").text
'			param.selectSingleNode("inputType").text = "hidden"
'			processParamField param
'			param.selectSingleNode("inputType").text = orgInputType		    		    
'		    end if
'		else
'		    	processParamField param		    
'	      	end if
'	    end if
'	next
'end if
%>
</TABLE>
<object data="../inc/calendar.asp" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
<%If formFunction = "query" then %>
        <%if (HTProgRight AND 2) <> 0 then %>
            <input type=button value ="查　詢" class="cbutton" onClick="formSearchSubmit()">
            <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <%end if%>    
<%Elseif formFunction = "edit" then %>
    <%if xShowType="5" then%>
       <% if (HTProgRight AND 16) <> 0 then %>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>               
    <%else%>

        <%if (HTProgRight AND 8) <> 0 then %>
               <input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()">
        <% End If %>
        
	<%if (HTProgRight and 16)<>0 AND _
		nullText(htPageDom.selectSingleNode("//sBaseDSDXML/addYN"))<>"N"	then%>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>           
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">
       <% if xShowType = "5" then %>
             <input type=button value ="編修參照來源" class="cbutton" onClick="EditRefParent()">   
       <%end If %>           
    <%end if%>
<% Else '-- add ---%>          
          <%if (HTProgRight AND 4)<>0 then %>
               <input type=button value ="新增存檔" class="cbutton" onClick="formAddSubmit()">
          <%end if%>     
          
          <%if (HTProgRight AND 4)<>0  then %>
              <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
          <%end if%>    

<% End If %>
</form>     
