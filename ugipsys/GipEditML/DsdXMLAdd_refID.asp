<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="新增"
HTUploadPath=session("Public")+"data/"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>選取引用資料</title>
<link href="css/popup.css" rel="stylesheet" type="text/css">
<link href="css/editor.css" rel="stylesheet" type="text/css">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="PopFormName">單元資料維護 - 選取引用資料條件</div>
<form action="DsdXMLAdd_refID2.asp" method="POST" id="PopForm">
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=CtNodeIDStr value="">
<object data="/inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
<table cellspacing="0">
  <tr>
    <td class="Label">主題單元點選</td>
    <td colspan=3><input type="button" class="InputButton" value="點選目錄樹" onClick="CtUnitTree()">
    </td>  
  </tr>
  <tr>
    <td class="Label">主題單元名稱</td>
    <td colspan=3><input name="htx_CtUnitName" type="text" value="" size="60" readonly class=rdonly></td>
  </tr>
  <tr>
    <td class="Label">標題</td>
    <td><input name="htx_Stitle" type="text" class="InputText" value="" size="20"></td>
    <td class="Label">關鍵詞</td>
    <td><input name="htx_xKeyword" type="text" class="InputText" value="" size="20"></td>
  </tr>
  <tr>
    <td class="Label">建立日期</td>
    <td colspan=3>   
    <input name="pcShowhtx_IDateS" type="text" class="InputText" size="10" readonly onclick="VBS: popCalendar 'htx_IDateS','htx_IDateE'">
      ～<input name="htx_IDateS" type=hidden><input name="htx_IDateE" type=hidden>
        <input name="pcShowhtx_IDateE" type="text" class="InputText" size="10" readonly onclick="VBS: popCalendar 'htx_IDateE',''">
        </td>
  </tr>
</table>
<hr>
<input type="button" class="InputButton" value="查詢清單" onClick="formSearchSubmit()">
<input type="button" class="InputButton" value="重填" onClick="resetForm()">
<input type="button" class="InputButton" value="回新增主畫面" onClick="Javascript:window.navigate('DsdXMLAdd.asp')">
</form>
</body>
</html>
<script language=vbs>
sub CtUnitTree()
	window.open "DsdXMLAdd_refID_CID.asp",null,"height=500,width=350,scrollbars=yes"
end sub

sub CtUnitID(xCtNodeIDStr,xCatNameStr)
	PopForm.CtNodeIDStr.value=xCtNodeIDStr
	PopForm.htx_CtUnitName.value=xCatNameStr
end sub

sub resetForm
	PopForm.reset()
end sub
  
sub formSearchSubmit()
        PopForm.submitTask.value="LIST"
        PopForm.Submit
end Sub

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
        document.all("pcShow"&CanTarget).value=d7date(document.all.CalendarTarget.value)
        if followCanTarget<>"" then
        	document.all(followCanTarget).value=document.all.CalendarTarget.value
        	document.all("pcShow"&followCanTarget).value=d7date(document.all.CalendarTarget.value)
        end if
	end if
end sub   
</script>

