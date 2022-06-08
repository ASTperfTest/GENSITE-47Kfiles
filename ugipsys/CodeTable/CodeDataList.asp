<%@ CodePage = 65001 %>
<%HTProgCap="代碼維護"
HTProgCode="Pn90M02"
HTProgPrefix="CodeData" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Set RSreg = Server.CreateObject("ADODB.RecordSet")
fSql=Request.QueryString("strSql")    
if fSql="" then
	fsql = "SELECT * FROM CodeMetaDef WHERE 1=1"
	for each x in request.form
	 if request(x) <> "" then 
	  if mid(x,2,3) = "fx_" then	  
		select case left(x,1)
		  case "s"
			fsql = fsql & " AND " & mid(x,5) & "=" & pkStr(request(x),"")
		  case else
			fsql = fsql & " AND " & mid(x,5) & " LIKE N'%" & request(x) & "%'"
		end select
	  end if
	 end if
	next
	fsql=fsql & " Order By CodeTblName"
end if

nowPage=Request.QueryString("nowPage")  '現在頁數


'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------

if RSreg.EOF then%>
	<script language=VBS>
		alert "找不到資料, 請重設查詢!"
		window.history.back
	</script>
<%	response.end
end if

totRec=RSreg.Recordcount       '總筆數

PerPageSize=cint(Request.QueryString("pagesize"))

if PerPageSize <= 0 then  
   PerPageSize=10  
end if 

RSreg.PageSize=PerPageSize       '每頁筆數

if cint(nowPage)<1 then 
   nowPage=1
elseif cint(nowPage) > RSreg.PageCount then 
   nowPage=RSreg.PageCount 
end if            	

RSreg.AbsolutePage=nowPage
totPage=RSreg.PageCount       '總頁數

strSql=server.URLEncode(fSql)  
%>
<HTML>
<head>
<title>查詢結果清單</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2"><%=Title%> </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" width="60%">【代碼資料查詢結果清單】</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><% if (HTProgRight and 2)=2 then %><a href="<%=HTProgPrefix%>Query.asp">重設查詢</a><% End IF %>&nbsp;</td>
  </tr>  
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>  
  <tr>
    <td width="80%" height=220 valign=top colspan="2">
    
<Form name=reg method="GET" action=<%=HTprogPrefix%>List.asp>
<center>
<font size="2" color="#0000FF">&nbsp;       
<%if not RSreg.eof  then%>                    
     <font size="2" color="rgb(63,142,186)"> 第
     <font size="2" color="#FF0000"><%=nowPage%>/<%=totPage%></font>頁|                      
        共<font size="2" color="#FF0000"><%=totRec%></font>
       <font size="2" color="rgb(63,142,186)">筆| 跳至第       
         <select id=GoPage size="1" style="color:#FF0000">
             <%For iPage=1 to totPage%> 
                   <option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
             <%Next%>   
         </select>      
    頁</font>    
        
       <% if nowPage <>1 then %>             
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">上一頁</a> 
       <%end if%>      
       
        <% if nowPage<>RSreg.PageCount then %> 
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">下一頁</a> 
        <%end if%>     
        | 每頁筆數:
       <select id=PerPage size="1" style="color:#FF0000">            
             <option value="10"<%if PerPageSize=10 then%> selected<%end if%>>10</option>                       
             <option value="20"<%if PerPageSize=20 then%> selected<%end if%>>20</option>
             <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
        </select>     
     </font>     
</center>
<CENTER>
<TABLE width=95% cellspacing="1" cellpadding="0" class=bg>                   
<tr align=left>    
	<td width=25% align=center class=lightbluetable>代碼ID</td>
	<td width=35% align=center class=lightbluetable>代碼名稱</td>
	<td width=40% align=center class=lightbluetable>存放資料表</td>
</tr>	                                
  <%                  
   for i=1 to PerPageSize%>                   
<tr>                  
	<TD class=whitetablebg><p align=center><font size=2><a href="CodeDataDetailList.asp?CodeID=<%=RSreg("CodeID")%>&CodeName=<%=RSreg("CodeName")%>"><%=RSreg("CodeID")%></A></FONT></TD>
	<TD class=whitetablebg><font size=2>　<%=RSreg("CodeName")%></FONT></TD>
	<TD class=whitetablebg><font size=2>　<%=RSreg("CodeTblName")%></FONT></TD>
</tr>
    <% RSreg.moveNext   
     if RSreg.EOF then exit for       
  next%>
    </td>
  </tr>  
  </table>
</table>
</CENTER>
</form>
</BODY>
</HTML>
<%else%>
<script language=vbs>
    msgbox "找不到資料, 請重設查詢條件!"
	window.history.back
</script>
<%end if%>
<script language=VBScript>	
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&strSql=<%= strSql%>" & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>" & "&strSql=<%= strSql%>" & "&pagesize=" & newPerPage                    
     end sub 
</script>
