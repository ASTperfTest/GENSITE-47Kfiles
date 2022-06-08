<%@ CodePage = 65001 %>
<% Response.Expires = 0 
HTProgCap="主題館績效統計"
response.charset = "utf-8"
HTProgCode = "webgeb2"
   %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
subjectid = request("subjectid")
session("subjects") = ""
session("DateS") = ""
session("DateE") = ""
response.write "<hr>"
%>
<html>
<head>
<title><%= HTProgCap %></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="http://www.mof.gov.tw/xslgip/mof94/CSS/lpcp.css" rel="stylesheet" type="text/css"/>
<script language="JavaScript" src="ts_picker.js"></script>
<link href="../css/list.css" rel="stylesheet" type="text/css">
<link href="../css/layout.css" rel="stylesheet" type="text/css">
<link href="../css/form.css" rel="stylesheet" type="text/css" />
</head>
<body leftmargin="0" rightmargin="0" topmargin="10" marginwidth="0" marginheight="0">
 <form name="form1" action="result.asp" method="post">
<object data="/inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
<CENTER><TABLE name="subjectlist" border="0" class="eTable" cellspacing="1" cellpadding="2" width="85%">
<TR><TD class="eTableLable" colspan=4 align="center">【<%=HTProgCap%>】</TR>
<TR><TD class="eTableLable" colspan="4" align="center">日期範圍：
<input name="pcShowhtx_IDateS" size="8" readonly onclick="VBS: popCalendar 'htx_IDateS','htx_IDateE'"> ～ 
<input name="htx_IDateS" type=hidden><input name="htx_IDateE" type=hidden>
<input name="pcShowhtx_IDateE" size="8" readonly onclick="VBS: popCalendar 'htx_IDateE',''">
</TD>
</TR>
<tr>
	<td colspan=4 align="center"> | <a href="subjectList.asp?subjectid=all">全部</a> | 
		<a href="subjectList.asp?subjectid=1">農</a> | 
		<a href="subjectList.asp?subjectid=5">林</a> | 
		<a href="subjectList.asp?subjectid=7">漁 </a>| 
		<a href="subjectList.asp?subjectid=8">牧</a> | 
		<a href="subjectList.asp?subjectid=9">其他</a> |<hr>
	</td>
</tr>
<%if subjectid = "all" then%>
<tr>
	<td align="right">農</td>
</tr>
<tr><TD class="eTableLable" align="right"></TD>
<%
sql = "SELECT CatTreeRoot.CtRootID ,CatTreeRoot.CtRootName"
sql = sql & " FROM NodeInfo LEFT OUTER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
sql = sql & " WHERE (NodeInfo.type1 = '農') and CatTreeRoot.inUse = 'Y'"
set rs = conn.execute(sql)
i=1
while not rs.eof
%>
<td><input type=checkbox name="ckbox" value=<%=rs("CtRootID")%>><%=rs("CtRootName")%></td>
<%
	if i mod 3 = 0 then
		response.write "</tr><tr><TD class='eTableLable' align='right'></TD>"
	end if
	i=i+1
	rs.movenext
wend
%>
</tr>
<tr>
	<td align="right">林</td>
</tr>
<tr><TD class="eTableLable" align="right"></TD>
<%
sql = "SELECT CatTreeRoot.CtRootID ,CatTreeRoot.CtRootName"
sql = sql & " FROM NodeInfo LEFT OUTER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
sql = sql & " WHERE     (NodeInfo.type1 = '林') and CatTreeRoot.inUse = 'Y'"
set rs = conn.execute(sql)
i=1
while not rs.eof
%>
<td><input type=checkbox name="ckbox" value=<%=rs("CtRootID")%>><%=rs("CtRootName")%></td>
<%
	if i mod 3 = 0 then
		response.write "</tr><tr><TD class='eTableLable' align='right'></TD>"
	end if
	i=i+1
	rs.movenext
wend
%>
</tr>
<tr>
	<td align="right">漁</td>
</tr>
<tr><TD class="eTableLable" align="right"></TD>
<%
sql = "SELECT CatTreeRoot.CtRootID ,CatTreeRoot.CtRootName"
sql = sql & " FROM NodeInfo LEFT OUTER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
sql = sql & " WHERE     (NodeInfo.type1 = '漁') and CatTreeRoot.inUse = 'Y'"
set rs = conn.execute(sql)
i=1
while not rs.eof
%>
<td><input type=checkbox name="ckbox" value=<%=rs("CtRootID")%>><%=rs("CtRootName")%></td>
<%
	if i mod 3 = 0 then
		response.write "</tr><tr><TD class='eTableLable' align='right'></TD>"
	end if
	i=i+1
	rs.movenext
wend
%>
</tr>
<tr>
	<td align="right">牧</td>
</tr>
<tr><TD class="eTableLable" align="right"></TD>
<%
sql = "SELECT CatTreeRoot.CtRootID ,CatTreeRoot.CtRootName"
sql = sql & " FROM NodeInfo LEFT OUTER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
sql = sql & " WHERE     (NodeInfo.type1 = '牧') and CatTreeRoot.inUse = 'Y'"
set rs = conn.execute(sql)
i=1
while not rs.eof
%>
<td><input type=checkbox name="ckbox" value=<%=rs("CtRootID")%>><%=rs("CtRootName")%></td>
<%
	if i mod 3 = 0 then
		response.write "</tr><tr><TD class='eTableLable' align='right'></TD>"
	end if
	i=i+1
	rs.movenext
wend
%>
</tr>
<tr>
	<td align="right">其他</td>
</tr>
<tr><TD class="eTableLable" align="right"></TD>
<%
sql = "SELECT CatTreeRoot.CtRootID ,CatTreeRoot.CtRootName"
sql = sql & " FROM NodeInfo LEFT OUTER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
sql = sql & " WHERE (NodeInfo.type1 = '其他') and CatTreeRoot.inUse = 'Y'"
set rs = conn.execute(sql)
i=1
while not rs.eof
%>
<td><input type=checkbox name="ckbox" value=<%=rs("CtRootID")%>><%=rs("CtRootName")%></td>
<%
	if i mod 3 = 0 then
		response.write "</tr><tr><TD class='eTableLable' align='right'></TD>"
	end if
	i=i+1
	rs.movenext
wend
%>
</tr>
<%else
sql = "SELECT CatTreeRoot.CtRootID ,CatTreeRoot.CtRootName ,NodeInfo.type1 "
sql = sql & " FROM  NodeInfo LEFT OUTER JOIN Type ON NodeInfo.type1 = Type.classname "
sql = sql & " LEFT OUTER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
sql = sql & " WHERE (Type.classid = '"& subjectid &"') and CatTreeRoot.inUse = 'Y'"
set rs = conn.execute(sql)
%>
<tr>
	<td align="right"><%=rs("type1")%></td>
</tr>
<tr><TD class="eTableLable" align="right"></TD>
<%
i=1
while not rs.eof
%>
<td><input type=checkbox name="ckbox" value=<%=rs("CtRootID")%>><%=rs("CtRootName")%></td>
<%
	if i mod 3 = 0 then
		response.write "</tr><tr><TD class='eTableLable' align='right'></TD>"
	end if
	i=i+1
	rs.movenext
wend
%>
</tr>
<%end if%>
<tr><td colspan=4 align=center>
	<hr>
	<input type="submit" class="cbutton" name="submit" value="送出">
	<input type="button" class="cbutton" name="all" value="全選" onClick="ChkAll">
	<input type="reset" class="cbutton" name="reset" value="重設">
	</td>
</tr>
</TABLE>
</CENTER>
</form>
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
sub Chkall
       chkCount=0     
       if document.all("all").value="全選" then           '全勾
          for i=0 to form1.elements.length-1
               set e=document.form1.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=true
                   chkCount=chkCount+1
              end if     
          next
       end if
   end sub   
</script>