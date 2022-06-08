<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border=0 class=bluetable cellspacing=1 cellpadding=2 width="80%">
		<TR>
		  <TD class=lightbluetable align=right>AP代碼</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=pfx_apcode SIZE=20<%IF request.querystring("APcode") <> "" then %> class=sedit readonly<% End If %>></TD>
		  <TD class=lightbluetable align=right>類別</TD>
		  <TD class=whitetablebg><Select name="sfx_apcat" size=1>
              					<option value="">請選擇</option>
 			<%SQL="Select apcatId,apcatCname from Apcat Order by apseq"
			SET RSS=conn.execute(SQL)
			if not RSS.EOF then
			While not RSS.EOF%>
			<option value="<%=RSS(0)%>"><%=RSS(0)%>--<%=RSS(1)%></option>
			<%	RSS.movenext
			wend
			end if%>
            	</select>
		  </TD>		  
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>英文名稱</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_apnameE SIZE=20></TD>
		  <TD class=lightbluetable align=right>中文名稱</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_apnameC SIZE=20></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>路徑</TD>
		  <TD class=whitetablebg colspan=3><INPUT TYPE=text NAME=tfx_appath SIZE=50></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>另開視窗</TD>
		  <TD class=whitetablebg><Select name="sfx_xsNewWindow" size=1>
              					<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId='boolYN' Order by msortValue DESC"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
            	</select>
		  </TD>		  
		  <TD class=lightbluetable align=right>提交ID/PW</TD>
		  <TD class=whitetablebg><Select name="sfx_xsSubmit" size=1>
              					<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId='boolYN' Order by msortValue DESC"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
            	</select>
		  </TD>		  
		</TR>
<%If formFunction <> "xquery" then %>
		<TR>
		  <TD class=lightbluetable align=right>顯示順序</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=tfx_aporder SIZE=20></TD>
		  <TD class=lightbluetable align=right>APMask</TD>
		  <TD class=whitetablebg><INPUT TYPE=text NAME=nfx_apmask SIZE=20 style="text-align:right;" onChange="VBS:initBitsArray me.value"></TD>
		</TR>
		<TR>
		  <TD class=lightbluetable align=right>權限 mask</TD>
		  <TD class=whitetablebg colspan="3" align="center">
        <table border="0" width="80%" cellspacing="1" class="bluetable" style="Display:'';" ID="M<%=cno%>" cellpadding="2">
          <tr class="lightbluetable">
            <td width="7%" align="center">查詢</td>
            <td width="7%" align="center">檢視</td>            
            <td width="7%" align="center">新增</td>
            <td width="7%" align="center">修改</td>
            <td width="7%" align="center">刪除</td>
            <td width="7%" align="center">列印</td>
            <td width="7%" align="center"><INPUT TYPE=text NAME=tfx_spare64 SIZE=6></td>            
            <td width="7%" align="center"><INPUT TYPE=text NAME=tfx_spare128 SIZE=6></td>
          </tr>
          <tr>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="maskBit2" value="Y" id="xmaskBit2" onClick="VBS:bitChange 2,2"></td>            
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="maskBit1" value="Y" id="xmaskBit1" onClick="VBS:bitChange 1,1"></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="maskBit3" value="Y" id="xmaskBit3" onClick="VBS:bitChange 3,4"></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="maskBit4" value="Y" id="xmaskBit4" onClick="VBS:bitChange 4,8"></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="maskBit5" value="Y" id="xmaskBit5" onClick="VBS:bitChange 5,16"></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="maskBit6" value="Y" id="xmaskBit6" onClick="VBS:bitChange 6,32"></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="maskBit7" value="Y" id="xmaskBit7" onClick="VBS:bitChange 7,64"></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="maskBit8" value="Y" id="xmaskBit8" onClick="VBS:bitChange 8,128"></td>
          </tr>
         </table>
		  </TD>
		</TR>
<% end if %>
</TABLE></CENTER>
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
<script language=vbs>
sub bitChange(xm, xValue)
	if document.all("maskBit"&xm).checked = true then
		reg.nfx_APMask.value = reg.nfx_APMask.value + xValue
	else
		reg.nfx_APMask.value = reg.nfx_APMask.value - xValue
	
	end if
end sub
    
    sub initBitsArray(value)
    	xp = 1
    	for xi = 1 to 8
    		if (value AND xp) = xp then
    			document.all("maskBit"&xi).checked = true
    		else
    			document.all("maskBit"&xi).checked = false
    		end if
    		xp = xp * 2
    	next
    end sub
</script>
