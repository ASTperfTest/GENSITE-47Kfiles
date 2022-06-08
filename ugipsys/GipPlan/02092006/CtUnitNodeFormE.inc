<form id="Form1" method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="CtUNitNodeEdit.asp?CtNodeID=<% =request.querystring("CtNodeID") %>">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE  cellspacing="0">
<TR><TD class="Label" align="right">
節點類別：</TD>
<TD class="whitetablebg">單元
<input type="hidden" name="htx_CtNodeKind">
<input type="hidden" name="htx_CtRootID">
<input type="hidden" name="htx_EditUserID">
<input type="hidden" name="htx_EditDate">
<input type="hidden" name="htx_CtNodeID">
</TD>
</TR>
<TR><TD class="Label" align="right"><font color="red">*</font>
標題：</TD>
<TD class="whitetablebg"><input name="htx_CatName" size="30">
</TD>
</TR>
<TR><TD class="Label" align="right">
主題單元：</TD>
<TD class="whitetablebg"><Select name="htx_ctUnitId" size=1>
<option value="">請選擇</option>
			<%SQL="Select ctUnitId,ctUnitName from CtUnit where ctUnitId IS NOT NULL AND inUse='Y'" _
				& " AND (deptId IS NULL OR deptId LIKE '" & session("deptId") & "%')  Order by ctUnitName"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="Label" align="right">功能類型：</TD>
<TD class="whitetablebg"><input name="dsp_refCtUnitKind" class="rdonly" readonly="true">
</TD>
</TR>
<TR><TD class="Label" align="right">單元資料範本：</TD>
<TD class="whitetablebg"><input name="dsp_refiBaseDSD" class="rdonly" readonly="true">
</TD>
</TR>
<TR><TD class="Label" align="right">標題圖示：</TD>
<TD class="whitetablebg"><table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr><td colspan="2">
		<input type="file" name="htImg_CtNameLogo">
		<input type="hidden" name="hoImg_CtNameLogo">
		<input type="hidden" name="htImgActCK_CtNameLogo">
	</td></tr>
	<tr>
	<td width="37%"><img id="logo_CtNameLogo" src="">
		<div id="noLogo_CtNameLogo" style="color:red">無圖片</div></td>
	<td valign="bottom">
		<div id="LbtnHide0_CtNameLogo">
		<% if (HTProgRight and 4)<>0 then %>
			<img id="addLogo_CtNameLogo" class="hand" src="../pagestyle/addimg.gif" onClick="VBS: addLogo 'CtNameLogo'">
		<% End if %>
		</div>
		<div id="LbtnHide1_CtNameLogo">
		<% if (HTProgRight and 8)<>0 then %>
			<img id="chgLogo_CtNameLogo" class="hand" src="../pagestyle/chimg.gif" onClick="VBS: chgLogo 'CtNameLogo'">
		<% End if %>
		<% if (HTProgRight and 16)<>0 then %>
			<img id="delLogo_CtNameLogo" class="hand" src="../pagestyle/delimg.gif" onClick="VBS: delLogo 'CtNameLogo'">
		<% End if %>
		</div>
		<div id="LbtnHide2_CtNameLogo">
		<% if (HTProgRight and 8)<>0 then %>
			<img id="orgLogo_CtNameLogo" class="hand" src="../pagestyle/resetimg.gif" onClick="VBS: orgLogo 'CtNameLogo'">
		<% End if %>
		</div>
	</td></tr></table>
</TD>
</TR>
<TR><TD class="Label" align="right">顯示次序：</TD>
<TD class="whitetablebg"><input name="htx_CatShowOrder" size="3">
</TD>
</TR>
<TR><TD class="Label" align="right"><font color="red">*</font>
是否開放：</TD>
<TD class="whitetablebg"><Select name="htx_inUse" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='bOnOff' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="Label" align="right">條列版型：</TD>
<TD class="whitetablebg"><Select name="htx_xslList" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='showClientStyle' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="Label" align="right">內容版型：</TD>
<TD class="whitetablebg"><Select name="htx_xslData" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='formClientStyle' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="Label" align="right">限制條件：</TD>
<TD class="whitetablebg"><input name="htx_dCondition" size="50">
</TD>
</TR>
<%if RSreg("CtNodeKind")="C" then%>
<TR><TD class="Label" align="right">
NP頁面類別：</TD>
<TD class="eTableContent"><Select name="htx_CtNodeNPKind" size=1>
<option value="">顯示所屬第一個子節點資料</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='CtNodeNPKind' Order by msortValue"
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
<TR><TD class="Label" align="right">
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

<!--TR><TD class="Label" align="right">
RSS：</TD>
<TD class="eTableContent"><input name="htx_RSS" size="50">
</TD>
</TR-->
<%end if%>
<%if checkGIPconfig("KMCat") then%>
<TR><TD class="Label" align="right">
KM分類：</TD>
<TD class="eTableContent">
<input name="htx_KMautoID" type="hidden">
<input name="htx_KMCatID" type="hidden">
<input name="htx_KMCat" size="40">
<input type=button value="外部分類" name="btn_KMcat" >
</TD>
</TR>
<TR><TD class="Label" align="right">
<%end if%>
<%if checkGIPconfig("SubjectCat") then%>
主題館ID：</TD>
<TD class="eTableContent"><input name="htx_SubjectID" size="50">
</TD>
</TR>
<%end if%>
<%if checkGIPconfig("SubjectMonth") then%>
<TR><TD class="Label" align="right">
主題館時令：</TD>
<TD class="eTableContent">
	<%SQL="Select mCode,mValue from codeMain where mSortValue IS NOT NULL AND codeMetaID='mouthID' Order by mSortValue"
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
<%if goFlag then%>
<TR><TD class="Label" align="right">
特殊節點型態：</TD>
<TD class="eTableContent"><Select name="htx_RSSNodeType" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from codeMain where mSortValue IS NOT NULL AND codeMetaID='RSSNodeType' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<%end if%>
<%if checkGIPconfig("KMCat") then%>
<!--TR><TD class="Label" align="right">
知識樹節點：</TD>
<TD class="eTableContent"><input type=checkbox name=bfx_KMTreeYN>
</TD>
</TR-->
<%end if%>
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
               <input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()" id="Save">
        <% End If %>
        
       <% if (HTProgRight AND 16) <> 0 then 
       		if not (RSreg("CtNodeKind")="C" and xchildCount > 0) then
       %>
       
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <% 	end if
       
          end If %>           
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
		