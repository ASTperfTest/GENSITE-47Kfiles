<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="CtUnitNodeAdd.asp?ItemID=<% =request.querystring("ItemID") %>&CatID=<% =request.querystring("CatID") %>&ctNodeKind=<% =request.querystring("ctNodeKind") %>">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="eTableLable" align="right">
節點類別：</TD>
<TD class="eTableContent"><%=RSC(0)%>
<input type="hidden" name="htx_CtNodeKind">
<input type="hidden" name="htx_CtRootID">
<input type="hidden" name="htx_DataLevel">
<input type="hidden" name="htx_dataParent">
<input type="hidden" name="htx_EditUserID">
<input type="hidden" name="htx_EditDate">
<input type="hidden" name="htx_CtNodeID">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
標題：</TD>
<TD class="eTableContent"><input name="htx_CatName" size="30">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">
主題單元：</TD>
<TD class="eTableContent"><Select name="htx_ctUnitId" size=1>
<option value="">請選擇</option>
			<%SQL="Select ctUnitId,ctUnitName from CtUnit where ctUnitId IS NOT NULL AND inUse='Y' Order by ctUnitName"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
&nbsp;&nbsp;
<BUTTON id="pickCtUnit" onClick="goAddUnit()">新增單元</BUTTON>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">標題圖示：</TD>
<TD class="eTableContent"><input type="file" name="htImg_CtNameLogo">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">顯示次序：</TD>
<TD class="eTableContent"><input name="htx_CatShowOrder" size="3">
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
是否開放：</TD>
<TD class="eTableContent"><Select name="htx_inUse" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId=N'bOnOff' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>

<%'======	2006.3.30 by Gary%>
<%if checkGIPconfig("RSSandQuery") then%>
<TR><TD class="eTableLable" align="right">
是否提供RSS：</TD>
<TD class="eTableContent"><Select name="htx_YNrss" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId=N'bOnOff' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">
是否提供查詢：</TD>
<TD class="eTableContent"><Select name="htx_YNquery" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId=N'bOnOff' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
</TD>
</TR>
<%end if%>
<%'======	%>
<%if request.querystring("CtNodeKind")="C" then%>
<TR><TD class="eTableLable" align="right">
NP頁面類別：</TD>
<TD class="eTableContent"><Select name="htx_CtNodeNPKind" size=1>
<option value="">顯示所屬第一個子節點資料</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId=N'CtNodeNPKind' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>

<%end if%>
<%if goFlag then%>
<TR><TD class="eTableLable" align="right">
RSS來源：</TD>
<TD class="eTableContent"><Select name="htx_RSSURLID" size=1>
<option value="">請選擇</option>
			<%SQL="Select RSSURLID,RSSURLName from RSSURL Where RSSURLShowYN='Y' Order by RSSURLSeq"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>

<!--TR><TD class="eTableLable" align="right">
RSS：</TD>
<TD class="eTableContent"><input name="htx_RSS" size="50">
</TD>
</TR-->
<TR><TD class="eTableLable" align="right">
KM分類：</TD>
<TD class="eTableContent">
<input name="htx_KMautoID" type="hidden">
<input name="htx_KMCatID" type="hidden">
<input name="htx_KMCat" size="40">
<input type=button value="外部分類" name="btn_KMcat" >
</TD>
</TR>
<TR><TD class="eTableLable" align="right">
主題館ID：</TD>
<TD class="eTableContent"><input name="htx_SubjectID" size="50">
</TD>
</TR>
<%if checkGIPconfig("SubjectMonth") then%>
<TR><TD class="eTableLable" align="right">
主題館時令：</TD>
<TD class="eTableContent">
	<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId=N'mouthID' Order by msortValue"
	SET RSS=conn.execute(SQL)
	mcount=0
	While not RSS.EOF
		mcount=mcount+1
	%>
	<input type=checkbox name="bfx_SubjectMonth" value="<%=RSS(0)%>"><%=RSS(1)%>&nbsp;&nbsp;&nbsp;
	<%	if mcount mod 6 = 0 then response.write "<br>"
		RSS.movenext
	wend%>
</TD>
</TR>
<%end if%>

<TR><TD class="eTableLable" align="right">
特殊節點型態：</TD>
<TD class="eTableContent"><Select name="htx_RSSNodeType" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId=N'RSSNodeType' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>

<!--TR><TD class="eTableLable" align="right">
知識樹節點：</TD>
<TD class="eTableContent"><input type=checkbox name=bfx_KMTreeYN>
</TD>
</TR-->
<%end if%>
<!--會員等級請加下列欄位-->
<%if checkGIPconfig("CtNodeMember") then%>
<TR><TD class="eTableLable" align="right">
會員類別等級：</TD>
<TD class="eTableContent">
	<%SQL="Select * from CtNodeMember order by mSortValue "
	SET RSS=conn.execute(SQL)
	mcount=0
	While not RSS.EOF
		mcount=mcount+1
	%>
		<input type=checkbox name="bfx_MType" value="<%=RSS(0)%>"><%=RSS(1)%>&nbsp;
		等級<Select name=sfx_MGrade<%=RSS(0)%> size=1>
	<%	
		for k=1 to RSS("grade")
			%><option value="<%=k%>"><%=k%></option><%
		next
		%></Select>&nbsp;&nbsp;&nbsp;<%
		if mcount mod 3 = 0 then response.write "<br>"
		RSS.movenext
	wend%>
</TD>
</TR>
<%end if%>
<!--會員等級-->
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
    
sub xaddOption(xlist,name,value)
	set xOption = document.createElement("OPTION")
	xOption.text=name
	xOption.value=value
	xlist.add(xOption)
end sub

sub xReturn(retValue)
	if retValue <> "" then
'		msgBox retValue
'		xPos = inStr(retValue, ";;;")
'		xEmpID = left(retValue,xPos-1)
'		xEmpName = mid(retValue,xPos+3)
'		msgBox xEmpID & "--" & xEmpName
		rtvArray = split(retValue,";;;")
		if uBound(rtvArray)<>1 then exit sub
	set xsrc = document.all("htx_ctUnitId")
	xaddOption xsrc, rtvArray(1), rtvArray(0)
'	xsrc.selectedIndex = xsrc.options.length-1
	xsrc.value = rtvArray(0)
	
'		reg.htx_Editor.value = rtvArray(0)
'		reg.htx_ChStuName.value = rtvArray(1)
	end if
end sub
</SCRIPT>
<SCRIPT LANGUAGE="javascript">
function goAddUnit() {
	window.open("/GipDSD/xCtUnitAdd.asp?phase=add&xn="+encodeURI(reg.htx_CatName.value),null,"height=400,width=600,status=no,toolbar=no,menubar=no,location=no");
}
</SCRIPT>
	


