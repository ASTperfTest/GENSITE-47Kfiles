<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="CuPageEdit.asp?ixCuPage=<% =request.querystring("ixCuPage") %>">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="eTableLable" align="right"><font color="red">*</font>
連結名稱：</TD>
<TD class="eTableContent"><input name="htx_aTitle" size="50">
<input type="hidden" name="htx_xiCuItem">
<input type="hidden" name="htx_ixCuPage">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">說明：</TD>
<TD class="eTableContent"><input name="htx_aDesc" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
連結頁：</TD>
<TD class="eTableContent"><input type="hidden" name="htx_nPageID">
<input name="show_nPageID" size="50" readonly class="rdonly">
<input type=button value="選文章" class="cbutton" onClick="pickPage()"
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
顯示於連結條列：</TD>
<TD class="eTableContent"><Select name="htx_bList" size=1>
<option value="Y">是</option>
<option value="N">否</option>
</select>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">條列次序：</TD>
<TD class="eTableContent"><input name="htx_listSeq" size="2">
</TD>
</TR>
</TABLE>
</CENTER>
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

    sub pickPage()
    	window.open "pickPageQuery.asp?wImg=Y&rtFunc=setPage"
    end sub
    
    sub setPage(rtValue)
		set oXML = createObject("Microsoft.XMLDOM")
  		oXml.loadXML rtValue
		if oXml.parseError.reason <> "" then 
			alert "傳回值錯誤!"
			exit sub
		end if  	
'		msgBox oXml.xml
		for each xr in oXml.selectNodes("//rtItem")
			reg.show_nPageID.value = xr.text
			reg.htx_nPageID.value = xr.getAttribute("id")
		next
    end sub
</SCRIPT>
	
		