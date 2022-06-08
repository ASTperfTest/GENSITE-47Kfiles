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

	       <a href="paActList.asp.asp">回條列</a>
	</div>
	<div id="ClearFloat"></div>
</div>

<form id="Form1" method="POST" action="docAttachAdd_act.asp" ENCTYPE="MULTIPART/FORM-DATA">
<object data="../../../inc/calendar1.htm" id="calendar1" type="text/x-scriptlet" width="245" height="160" style="position:absolute;top:0;left:230;visibility:hidden">日曆選單</object>
<INPUT TYPE=hidden name=CalendarTarget>
<INPUT TYPE=hidden name="pid" value="<% =request("pid") %>">
<table cellspacing="0">

<TR><TD class="Label" align="right">檔案下載</TD><TD class="eTableContent"><input type="file" name="filename1">

	
</TD></TR>
<TR>
<TD class="Label" align="right">檔案說明</TD>
<TD class="eTableContent"><input name="filename1_note" size="50"></TD>
</TR>
</TABLE>

               <input type="submit" value ="確定" class="cbutton" >
        
            
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">
       
</form>     

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
