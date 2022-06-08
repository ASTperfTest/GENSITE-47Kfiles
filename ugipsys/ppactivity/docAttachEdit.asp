<% Response.Expires = 0
HTProgCap="課程"
HTProgFunc="編修"
HTProgCode="PA001"
HTProgPrefix="paAct" %>
<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<!--#include virtual = "/inc/hyftdGIP.inc" -->
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title>講義下載</title>
</head>

<body>
<div id="FuncName">
	<h1>講義下載</h1>
	<div id="Nav">
	       <a href="paActList.asp">回條列</a>
	</div>
	<div id="ClearFloat"></div>
</div>
<%
   sql="select * from docdownload_attach where id=" & request("id")
   set rs=conn.Execute(sql)
   if not rs.eof then
   filename1=trim(rs("filename"))
   filename1_note=trim(rs("filename_note"))

%>
<form id="Form1" method="POST" action="docAttachEdit_act.asp" ENCTYPE="MULTIPART/FORM-DATA">
<input type="hidden" name="id" value="<% =request("id") %>">
<input type="hidden" name="pid" value="<% =rs("pid") %>">
<object data="../../../inc/calendar1.htm" id="calendar1" type="text/x-scriptlet" width="245" height="160" style="position:absolute;top:0;left:230;visibility:hidden">日曆選單</object>
<INPUT TYPE=hidden name=CalendarTarget>
<table cellspacing="0">

<TR><TD class="Label" align="right">檔案下載</TD><TD class="eTableContent"><table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr><td colspan="2">
		<input type="file" name="filename1">
	</td></tr>
	<tr>
	<td width="37%"><span id="logo_fileDownLoad" class="rdonly"></span>
		<div id="noLogo_fileDownLoad" style="color:red"><% =filename1 %><input type="hidden" name="ofilename1" value="<% =filename1 %>"></div></td>
	</tr></table>

</TD></TR>
<TR>
<TD class="Label" align="right">檔案說明</TD>
<TD class="eTableContent"><input name="filename1_note" size="50"  value="<% =filename1_note %>"></TD>
</TR>
</TABLE>

               <input type="submit" name="submit" value ="修改" class="cbutton" >
               <input type="submit" name="submit" value ="刪除" class="cbutton"  onclick="return confirm('確定刪除此筆資料嗎?')">
            
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">
       
</form>     
<% 
   end if 
%>   
<script language="vbs">
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


     
</body>
</html>
