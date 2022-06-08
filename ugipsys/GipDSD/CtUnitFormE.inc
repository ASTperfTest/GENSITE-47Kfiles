<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="CtUnitEdit.asp?ctUnitId=<% =request.querystring("ctUnitId") %>">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="eTableLable" align="right">主題單元名稱：</TD>
<TD class="eTableContent"><input name="htx_CtUnitName" size="30" title="作為分類節點的預設值">
<input type="hidden" name="htx_CtUnitID">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">單元LOGO：</TD>
<TD class="eTableContent"><table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr><td colspan="2">
		<input type="file" name="htImg_CtUnitLogo">
		<input type="hidden" name="hoImg_CtUnitLogo">
		<input type="hidden" name="htImgActCK_CtUnitLogo">
	</td></tr>
	<tr>
	<td width="37%"><img id="logo_CtUnitLogo" src="">
		<div id="noLogo_CtUnitLogo" style="color:red">無圖片</div></td>
	<td valign="bottom">
		<div id="LbtnHide0_CtUnitLogo">
		<% if (HTProgRight and 4)<>0 then %>
			<img id="addLogo_CtUnitLogo" class="hand" src="../pagestyle/addimg.gif" onClick="VBS: addLogo 'CtUnitLogo'">
		<% End if %>
		</div>
		<div id="LbtnHide1_CtUnitLogo">
		<% if (HTProgRight and 8)<>0 then %>
			<img id="chgLogo_CtUnitLogo" class="hand" src="../pagestyle/chimg.gif" onClick="VBS: chgLogo 'CtUnitLogo'">
		<% End if %>
		<% if (HTProgRight and 16)<>0 then %>
			<img id="delLogo_CtUnitLogo" class="hand" src="../pagestyle/delimg.gif" onClick="VBS: delLogo 'CtUnitLogo'">
		<% End if %>
		</div>
		<div id="LbtnHide2_CtUnitLogo">
		<% if (HTProgRight and 8)<>0 then %>
			<img id="orgLogo_CtUnitLogo" class="hand" src="../pagestyle/resetimg.gif" onClick="VBS: orgLogo 'CtUnitLogo'">
		<% End if %>
		</div>
	</td></tr></table>
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
功能類型：</TD>
<TD class="eTableContent"><Select name="htx_CtUnitKind" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMainLong where msortValue IS NOT NULL AND codeMetaId='refCTUKind' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right">重導URL：</TD>
<TD class="eTableContent"><input name="htx_redirectURL" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">開新視窗：</TD>
<TD class="eTableContent"><Select name="htx_newWindow" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='boolYN' Order by msortValue"

			SET RSS=conn.execute(SQL)

			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right">單元資料範本：</TD>
<TD class="eTableContent">
<% if xCount>=1 then%>
<input type=hidden name="htx_iBaseDSD">
<input type=hidden name=xrefDSDCat value="">
<input type=hidden name=iBaseDSDSelect value="">
<%=RSreg("sbaseDsdname")%>&nbsp;&nbsp;
    <%if checkGIPconfig("MMOFolder") and RSreg("rdsdcat")="MMO" then 
    	SQLM="Select Case MM.mmofolderName when '/' then MS.mmositeName Else MM.mmofolderNameShow END MMOFolderName " & _
    		"from Mmofolder MM Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId where mmofolderId="&RSreg("mmofolderId")

			Set RSM=conn.execute(SQLM)
    	%>
    	<span id="MMOFolder"><input type=hidden name=htx_MMOFolderID value="">
    	(多媒體型主題單元存放根目錄:<%=RSM("mmofolderName")%>)</span>
    <%end if%>
<%else%>
<input type=hidden name=htx_iBaseDSD value=""><input type=hidden name=xrefDSDCat value="">
<Select name="iBaseDSDSelect" size=1 onChange="VBS:iBaseDSDSpilt()">
<option value="">請選擇</option>
			<%SQL="Select ibaseDsd,sbaseDsdname,rdsdcat from BaseDsd where ibaseDsd IS NOT NULL AND inUse='Y' Order by ibaseDsd"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>---<%=RSS(2)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
    <%if checkGIPconfig("MMOFolder") and RSreg("rdsdcat")="MMO" then %>
	<span id="MMOFolder">多媒體型主題單元存放根目錄
		<Select Name="htx_MMOFolderID" size=1>
		  <%
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
			sqlCom=sqlCom & " order by Case mmofolderParent when 'zzz' then '' else mmofolderParent END, mmofolderID"
'				response.write sqlCom
			set RSS = conn.execute(sqlCom)
			if not RSS.EOF then
				ARYDept = RSS.getrows(300)
				glastmsglevel = 0
				genlist 0, 0, 1, 0
			        expandfrom ARYDept(cid, 0), 0, 0
			end if
		  %>
		  </select>
		  <input type=button value ="新增子目錄" class="cbutton" onClick="MMOFolderAdd()">
	</span>
    <%end if%>
<%end if%>
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
只顯示此單元資料：</TD>
<TD class="eTableContent"><Select name="htx_fCtUnitOnly" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='boolYN' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
是否需要審稿：</TD>
<TD class="eTableContent"><Select name="htx_CheckYN" size=1>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='boolYN' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
</TD>
</TR>
<% 	if checkGIPconfig("CtUnitExpireCheck") then %>
<TR><TD class="eTableLable" align="right">單元資料逾期天數：</TD>
<TD class="eTableContent"><input name="htx_CtUnitExpireDay" size="3"> 天
</TD></TR>
<%	end if %>
<TR><TD class="eTableLable" align="right">單 位：</TD>
<TD class="eTableContent"><Select name="htx_deptId" size=1>
<option value="">--- 共 用 ---</option>
<%
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
%>
</select>
</TD></TR>
<TR><TD class="eTableLable" align="right">頁頭網頁ID：</TD>
<TD class="eTableContent"><input name="htx_HeaderPart" size="6">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">頁尾網頁ID：</TD>
<TD class="eTableContent"><input name="htx_FooterPart" size="6">
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
        
       <% if xCount<1 and (HTProgRight AND 16) <> 0 then %>
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
		
