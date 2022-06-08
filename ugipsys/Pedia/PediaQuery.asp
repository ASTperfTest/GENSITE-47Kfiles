<%@ CodePage = 65001 %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="資料清單"
HTProgCode="PEDIA01"
HTProgPrefix="PediaQuery" 

Dim iCtUnitId : iCtUnitId = request.querystring("ictunit")

%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title></title>
</head>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
	  <td width="50%" align="left" nowrap class="FormName">單元查詢&nbsp;</td>
		<td width="50%" class="FormLink" align="right">
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A>			
		</td>	
  </tr>
	<tr>
	  <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
	</tr>
	<tr>
		<td class="Formtext" colspan="2" height="15"></td>
	</tr>  
	<tr>
		<td align=center colspan=2 width=80% height=230 valign=top>    
		<form method="POST" name="reg" action="">
		<INPUT TYPE="hidden" name="submitTask" value="QUERY">
		<object data="/inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
		<INPUT TYPE="hidden" name=CalendarTarget>
		<CENTER>
		<TABLE border="0" id="ListTable" cellspacing="1" cellpadding="2">
		<% if iCtUnitId = session("PediaUnitId") then %>
		<TR>
			<Th>詞目</Th>
			<TD class="eTableContent"><input name="sTitle" size="50"></TD>
		</TR>
		<% end if %>
		<% if iCtUnitId = session("PediaAdditionalUnitId") then %>
		<TR>
			<Th>補充解釋</Th>
			<TD class="eTableContent"><input name="xBody" size="50"></TD>
		</TR>
		<% end if %>
		<TR>
			<Th>張貼日</Th>
			<TD class="eTableContent">
				<input name="xPostDateS" size="10" readonly onclick="VBS: popCalendar 'xPostDateS','xPostDateE'"> ～ 				
				<input name="xPostDateE" size="10" readonly onclick="VBS: popCalendar 'xPostDateE',''">
			</TD>
		</TR>
		<TR>
			<Th>審核狀態</Th>
			<TD class="eTableContent">
				<Select name="fCTUPublic">
					<option value="">請選擇</option>
					<option value="Y">通過</option>
					<option value="N">不通過</option>
				</select>
			</TD>
		</TR>
		<TR>
			<Th>發表會員</Th>
			<TD class="eTableContent"><input name="memberId" size="50"></TD>
		</TR>
		<TR>
			<Th>狀態</Th>
			<TD class="eTableContent">
				<Select name="xStatus">
					<option value="">請選擇</option>
					<option value="Y">未刪除</option>
					<option value="D">已刪除</option>
				</select>
			</TD>
		</TR>
		</TABLE>
		</td>
	</tr>
	</table>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td width="100%">     
			<p align="center">
				<input name="button" type="button" class="cbutton" onclick="formSubmit()" value ="查 詢">
        <input type="button" value ="重　填" class="cbutton" onClick="resetForm()">        
		</td>
	</tr>
	</table>
</body>
</html>
<script language=vbs>
	cvbCRLF = vbCRLF
	cTabchar = chr(9)

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
			'document.all("x"&CanTarget).value=document.all.CalendarTarget.value
			if followCanTarget<>"" then
				document.all(followCanTarget).value=document.all.CalendarTarget.value
				'document.all("x"&followCanTarget).value=document.all.CalendarTarget.value
			end if
		end if
	end sub   
 	
	sub formSubmit()
	
		<% if iCtUnitId = session("PediaUnitId") then %>
			reg.action = "PediaList.asp"
		<% end if %>
		<% if iCtUnitId = session("PediaAdditionalUnitId") then %>
			reg.action = "PediaAdditionalList.asp?icuitem=<%=request.querystring("picuitem")%>"
		<% end if %>
		reg.submit
		
	end sub
	
	sub resetForm 
		window.location.href = "/Pedia/PediaQuery.asp?ictunit=<%=request.querystring("ictunit")%>"
  end sub
</script>
