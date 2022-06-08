<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="點閱統計"
HTProgFunc="點閱統計"
HTProgCode="GW1M22"
HTProgPrefix="mSession" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
 Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>點閱統計</title>
</head>

<body>

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

Sub fmQuerySubmit()	
    
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
	
	IF stat1.htx_dbDate.value = Empty Then 
		MsgBox replace(nMsg,"{0}","資料範圍起日"), 64, "Sorry!"
		stat1.htx_dbDate.focus
		exit sub
	END IF
	IF (stat1.htx_dbDate.value <> "") AND (NOT isDate(stat1.htx_dbDate.value)) Then
		MsgBox replace(dMsg,"{0}","資料範圍起日"), 64, "Sorry!"
		stat1.htx_dbDate.focus
		exit sub
	END IF
	IF stat1.htx_deDate.value = Empty Then 
		MsgBox replace(nMsg,"{0}","資料範圍迄日"), 64, "Sorry!"
		stat1.htx_deDate.focus
		exit sub
	END IF
	IF (stat1.htx_deDate.value <> "") AND (NOT isDate(stat1.htx_deDate.value)) Then
		MsgBox replace(dMsg,"{0}","資料範圍迄日"), 64, "Sorry!"
		stat1.htx_deDate.focus
		exit sub
	END IF
  
  stat1.Submit
End Sub
</script>

<p>網站類別點閱統計</p>
<form name="stat1" action="type_stat.asp" method="post">
<object data="../inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>


資料範圍起日：<input name="htx_dbDate" type=hidden>
<input name="pcShowhtx_dbDate" size="8" readonly onclick="VBS: popCalendar 'htx_dbDate' ,''">

資料範圍迄日：<input name="htx_deDate" type=hidden>
<input name="pcShowhtx_deDate" size="8" readonly onclick="VBS: popCalendar 'htx_deDate' ,''">


<input type="button" value="開始統計" name="B3" onclick="fmQuerySubmit()"></p>
</form>
<%
  date1 = request("htx_dbDate")
  date2 = request("htx_deDate")
  
  'response.write "date1= "& date1 &"<BR>date2= "& date2 &"<BR>"

  if cint(isdate(date1))<>0 and cint(isdate(date2))<>0 then
  sql="select iCtUnit, max(n.CatName), count(*) as hitNumber from gipHitUnit AS u LEFT JOIN CatTreeNode AS n ON u.iCtnode=n.CtNodeID where hitTime between '" & date1 & "' AND '" & date2 & "' group by iCtUnit order by count(*) DESC"
  
  response.write "<!--sql= "& sql &"-->"
  'response.end
  
  set rs=conn.Execute(sql)
%>
<table border="1" width="100%">
  <tr>
    <td width="23%">類別ID</td>
    <td width="26%">類別名稱</td>
    <td width="74%">點閱次數</td>
  </tr>
  <% do while not rs.eof  %> 
  <tr>
    <td width="23%"><% =rs(0) %>&nbsp;</td>
    <td width="26%"><% =rs(1) %>&nbsp;</td>
    <td width="74%"><% =rs(2) %>&nbsp;</td>
  </tr>
<%  rs.MoveNext
    loop
%>   
</table>
<p><input type="button" value="轉成Excel" name="B3"></p>
<%
  end if
%>
</body>

</html>
