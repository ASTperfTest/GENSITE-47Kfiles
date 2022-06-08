<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="">
<INPUT TYPE=hidden name=submitTask value="">
<input type="hidden" name="id" value="<% =request("actID") %>">
<%
  check1_1="0"
  check1_2="0"
  check1_3="0"
  check1_4="0"
  check1_5="0"
  check1_6="0"
  check1_7="0"
  check1_8="0"
  check1_9="0"
  check1_10="0"
  check1_11="0"
  check1_12="0"
  check1_13="0"
  check1_14="0"
  check1_15="0"
  check1_16="0"
  check1_17="0"
  check1_18="0"
  check1_19="0"
  check1_20="0"
  check2_1="0"
  check2_2="0"
  check2_3="0"
  check2_4="0"
  check2_5="0"
  check2_6="0"
  check2_7="0"
  check2_8="0"
  check2_9="0"
  check2_10="0"
  check2_11="0"
  check2_12="0"
  check2_13="0"
  check2_14="0"
  check2_15="0"
  check2_16="0"
  check2_17="0"
  check2_18="0"
  check2_19="0"
  check2_20="0"
  
  sql01="select * from paInfoDesign where id=" & request("actID") 
  set rs01=conn.Execute(sql01)
  if not rs01.eof then
    check1_1=rs01("check1_1")
  check1_2=rs01("check1_2")
  check1_3=rs01("check1_3")
  check1_4=rs01("check1_4")
  check1_5=rs01("check1_5")
  check1_6=rs01("check1_6")
  check1_7=rs01("check1_7")
  check1_8=rs01("check1_8")
  check1_9=rs01("check1_9")
  check1_10=rs01("check1_10")
  check1_11=rs01("check1_11")
  check1_12=rs01("check1_12")
  check1_13=rs01("check1_13")
  check1_14=rs01("check1_14")
  check1_15=rs01("check1_15")
  check1_16=rs01("check1_16")
  check1_17=rs01("check1_17")
  check1_18=rs01("check1_18")
  check1_19=rs01("check1_19")
  check1_20=rs01("check1_20")
  check2_1=rs01("check2_1")
  check2_2=rs01("check2_2")
  check2_3=rs01("check2_3")
  check2_4=rs01("check2_4")
  check2_5=rs01("check2_5")
  check2_6=rs01("check2_6")
  check2_7=rs01("check2_7")
  check2_8=rs01("check2_8")
  check2_9=rs01("check2_9")
  check2_10=rs01("check2_10")
  check2_11=rs01("check2_11")
  check2_12=rs01("check2_12")
  check2_13=rs01("check2_13")
  check2_14=rs01("check2_14")
  check2_15=rs01("check2_15")
  check2_16=rs01("check2_16")
  check2_17=rs01("check2_17")
  check2_18=rs01("check2_18")
  check2_19=rs01("check2_19")
  check2_20=rs01("check2_20")
  end if
%>
<CENTER>
<TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="70%">
<tr>
<td  class="eTableLable" align="right">欄位名稱</td><td  class="eTableLable" align="left">前台顯示</td><td  class="eTableLable" align="left">是否必填</td>
</tr>
<TR><TD class="eTableLable" align="right">
登入密碼：</TD>

<TD class="eTableContent"><input type="checkbox" name="check1_1" value="1" <% if check1_1="1" then response.write " checked" end if %> ></td>
<TD class="eTableContent"><input type="checkbox" name="check2_1" value="1" <% if check2_1="1" then response.write " checked" end if %>></td>
</TR>
<TR>
<TD class="eTableLable" align="right"> 密碼確認：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_2" value="1" <% if check1_2="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_2" value="1" <% if check2_2="1" then response.write " checked" end if %>></td>
</TR>
<TR><TD class="eTableLable" align="right">
身份證號：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_3" value="1" <% if check1_3="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_3" value="1" <% if check2_3="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">
姓名：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_4" value="1" <% if check1_4="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_4" value="1" <% if check2_4="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">
出生日：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_5" value="1" <% if check1_5="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_5" value="1" <% if check2_5="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">性別：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_6" value="1" <% if check1_6="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_6" value="1" <% if check2_6="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">
家中電話：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_7" value="1" <% if check1_7="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_7" value="1" <% if check2_7="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">
行動電話：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_8" value="1" <% if check1_8="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_8" value="1" <% if check2_8="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">公司電話：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_9" value="1" <% if check1_9="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_9" value="1" <% if check2_9="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right"> 傳真：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_10" value="1" <% if check1_10="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_10" value="1" <% if check2_10="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">E-mail：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_11" value="1" <% if check1_11="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_11" value="1" <% if check2_11="1" then response.write " checked" end if %>></td>
</TR>
<TR><TD class="eTableLable" align="right">最高學歷：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_12" value="1" <% if check1_12="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_12" value="1" <% if check2_12="1" then response.write " checked" end if %>></td>
</TR>
<TR><TD class="eTableLable" align="right">服務機構：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_13" value="1" <% if check1_13="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_13" value="1" <% if check2_13="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">任職部門：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_14" value="1" <% if check1_14="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_14" value="1" <% if check2_14="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">職稱：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_15" value="1" <% if check1_15="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_15" value="1" <% if check2_15="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">聯絡地址：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_16" value="1" <% if check1_16="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_16" value="1" <% if check2_16="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">辦公地址：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_17" value="1" <% if check1_17="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_17" value="1" <% if check2_17="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">公司統一般號：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_18" value="1" <% if check1_18="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_18" value="1" <% if check2_18="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">消息來源：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_19" value="1" <% if check1_19="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_19" value="1" <% if check2_19="1" then response.write " checked" end if %>></td>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">飲食習慣：</TD>
<TD class="eTableContent"><input type="checkbox" name="check1_20" value="1" <% if check1_20="1" then response.write " checked" end if %>></td>
<TD class="eTableContent"><input type="checkbox" name="check2_20" value="1" <% if check2_20="1" then response.write " checked" end if %>></td>
</TD>
</TR>
</TABLE>
</CENTER>
<object data="../inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">

               <input type=button value ="編修存檔" class="cbutton" onClick="formAddSubmit()">
      
              <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
      
 </td></tr>
</table>
<SCRIPT LANGUAGE="vbs">
    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub
</SCRIPT>
	

<SCRIPT LANGUAGE="vbs">

</SCRIPT>
		<script language=vbs>
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
