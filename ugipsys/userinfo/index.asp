<%@ CodePage = 65001 %>
<%
  if request("Submit")="匯出LOG檔" then
   response.redirect "export.asp?select0=" & request("select0") & "&htx_CuDTx23F84S=" & request("htx_CuDTx23F84S") & "&htx_CuDTx23F84E=" & request("htx_CuDTx23F84E")
   response.end
end if
%>
<% Response.Expires = 0
HTProgCap="使用者登入系統記錄留存備查 "
HTProgFunc="使用者登入訊息"
HTProgCode="BM010"
HTProgPrefix="mSession" %>
<% response.expires = 0 %>
<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->

<!--#include FILE = "../inc/selectTree.inc" -->
<!--#include FILE = "../HTSDSystem/HT2CodeGen/htUIGen.inc" -->

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
</head>

<body>
<object data="../inc/calendar.htm" id="calendar1" type="text/x-scriptlet" width="245" height="160" style="position:absolute;top:0;left:230;visibility:hidden"></object>
<table border="0" width="95%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
</table>  
<%

queryuser=request("select0")

if queryuser<>"全部使用者" then
if request("htx_CuDTx23F84S")<>"" and request("htx_CuDTx23F84E")<>"" then
sql2="exec sp_userinfo N'" & queryuser & "',N'" & request("htx_CuDTx23F84S") & "',N'" & dateadd("d",1,request("htx_CuDTx23F84E")) & "'"
else
sql2="exec sp_userinfo N'" & queryuser & "','',''"
end if
else
if request("htx_CuDTx23F84S")<>"" and request("htx_CuDTx23F84E")<>"" then
sql2="exec sp_userinfo '','" & request("htx_CuDTx23F84S") & "','" & dateadd("d",1,request("htx_CuDTx23F84E")) & "'"
else
sql2="exec sp_userinfo '','',''"
end if
end if


set rs=conn.Execute(sql2)

%>
<form action="index.asp" method="post" name="reg" >
<INPUT TYPE="hidden" name="CalendarTarget" />
<p>使用者 
<select name="select0">
<option>全部使用者</option>
<% sql="select * from infouser"
   set rs2=conn.Execute(sql)
   do while not rs2.eof 
%>   
<option <% if trim(rs2("userid"))=request("select0") then response.write "selected" end if %>><% =trim(rs2("userid")) %>
</option>
<% 
  rs2.MoveNext
  loop
%>  
</select>
日期範圍 : <input name="pcShowhtx_CuDTx23F84S" size="10" readonly="true" onClick="VBS: popCalendar 'htx_CuDTx23F84S','htx_CuDTx23F84E'" onkeypress="VBS: popCalendar 'htx_CuDTx23F84S','htx_CuDTx23F84E'" /> ～ 
<input name="htx_CuDTx23F84S" type="hidden" /><input name="htx_CuDTx23F84E" type="hidden" />
<input name="pcShowhtx_CuDTx23F84E" size="10" readonly="true" onClick="VBS: popCalendar 'htx_CuDTx23F84E',''" onkeypress="VBS: popCalendar 'htx_CuDTx23F84E',''"/>
<input type="submit" name="submit" value="確定">
<input type="submit" name="Submit" value="匯出LOG檔" />
</form>
</p>

    <TABLE width=90% cellspacing="1" cellpadding="8" class=bg bgcolor=navy>                   
    <tr bgcolor="#CCCCCC">
    <td>姓名</td>
    <td>帳號</td>
    <td>登入IP</td>
    <td>時間</td>
    <td>狀態</td>
    <td>主題單元</td>
    <td>異動資料</td>
  </tr>
 <%
    
    While Not rs.Eof
   
    %>  
   <tr bgcolor=white> 
      <TD class=whitetablebg align=center><% =rs("UserName") %></td>
      <TD class=whitetablebg align=center><% =rs("userid") %></td>
      <TD class=whitetablebg align=center><% =rs("loginip") %></td>
      <TD class=whitetablebg align=center><% =rs("logintime") %></td>
      <TD class=whitetablebg align=center><% =rs("xtarget") %><% =rs("act") %></td>
      <TD class=whitetablebg align=center><% =rs("ctunit") %></td>
      <TD class=whitetablebg align=center><% =rs("objtitle") %></td>
    </tr>
<%        
     rs.MoveNext
    WEnd    
%>    
  </form>
</table>
<p><a href="javascript:history.go(-1);"><b>回上頁</b></a> </p>
</body>
</html>
    
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