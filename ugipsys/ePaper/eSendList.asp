<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="電子報管理"
HTProgFunc="發送清單"
HTProgCode="GW1M51"
HTProgPrefix="eSend" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
	Server.ScriptTimeOut = 5000
 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 'fSql=Request.QueryString("strSql")
 '新增查條件
if request("account") <> "" then
	session("account") = request("account")
else 
	session("account") = ""
end if
if request("realname") <> "" then
	session("realname") = request("realname")
else 
	session("realname") = ""
end if
if request("eMail") <> "" then
	session("eMail") = request("eMail")
else 
	session("eMail") = ""
end if
'----------added by Leo----------Start----------利用repeat來判斷是否檢視重複的eMail
if request("repeat") <> "" then
	session("repeat") = request("repeat")
else
	session("repeat") = ""
end if
'----------added by Leo-----------End-----------
	epubid = request("ePubID")
	repeat = request("repeat")
	

	fSql = "SELECT htx.*, m.account, m.realname "
	fSql = fSql	& ", (SELECT count(*) FROM memEPaper WHERE memEPaper.memID=m.account) AS memEPCount "
	fSql = fSql	& " FROM EpSend AS htx LEFT JOIN Member AS m ON htx.email=m.email"
	fSql = fSql	& " WHERE htx.ePubID=" & pkstr(request("ePubID"),"")
	
	'----------added by Leo----------Start----------
	if repeat = "1" then
		fSql = fSql & " AND htx.email IN"
		fSql = fSql & " (SELECT htx.email FROM EpSend AS htx LEFT OUTER JOIN Member AS m ON htx.email = m.email"
		fSql = fSql & " WHERE (htx.ePubID = " & pkstr(request("ePubID"),"") & ")"
		fSql = fSql & " GROUP BY htx.ePubID, htx.email, htx.sendDate, htx.CtRootID"
		fSql = fSql & " HAVING (COUNT(htx.email) > 1))"	
	end if		
	'----------added by Leo-----------End-----------
	
	'加入判斷式
	if session("account") <> "" then
		fSql = fSql & " and (m.account is not null and m.account like '%"& session("account") &"%') "
	end if
	if session("realname") <> "" then
		fSql = fSql & " and (m.realname is not null and (m.realname like '%"& session("realname") &"%' or m.realname like '%"& Chg_UNI(session("realname")) &"%' )) "
	end if
	if session("eMail") <> "" then
		fSql = fSql & " and (htx.eMail like '%"& session("eMail") &"%') "
	end if
	fSql = fSql	& " ORDER BY htx.email"

nowPage=Request.QueryString("nowPage")  '現在頁數

session("epaper_sql") = fSql
'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=50  
      end if 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
      'strSql=server.URLEncode(fSql)
   end if    
end if   

Function Chg_UNI(str)        'ASCII轉Unicode
	dim old,new_w,iStr
	old = str
	new_w = ""
	for iStr = 1 to len(str)
		if ascw(mid(old,iStr,1)) < 0 then
			new_w = new_w & "&#" & ascw(mid(old,iStr,1))+65536 & ";"
		elseif        ascw(mid(old,iStr,1))>0 and ascw(mid(old,iStr,1))<127 then
			new_w = new_w & mid(old,iStr,1)
		else
			new_w = new_w & "&#" & ascw(mid(old,iStr,1)) & ";"
		end if
	next
	Chg_UNI=new_w
End Function

%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
		<!-- Modify By Leo  新增判斷式(發送清單、重複清單)  --- Start --->
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【 <% If repeat = "1" then HTProgFunc = "重複清單" end if %> <%=HTProgFunc%> 】</td>
		<!-- Modify By Leo  新增判斷式(發送清單、重複清單)  ---  End  --->
		<td width="50%" class="FormLink" align="right">
			<A href="ePubList.asp" title="回電子報條列">回電子報條列</A> 
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
 <Form name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
  <p align="center">  
<%If not RSreg.eof then%>     
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
       <% if cint(nowPage) <>1 then %>
		<!--Modify by Leo  新增判斷式  -- Start -->
			<% if (repeat) <> "1" then %>
            |<a href="<%=HTProgPrefix%>List.asp?epubid=<%=epubid%>&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>">上一頁</a> 
			<% else %>
			|<a href="<%=HTProgPrefix%>List.asp?epubid=<%=epubid%>&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>&repeat=1">上一頁</a>
			<% end if %>
		<!--Modify by Leo  新增判斷式  -- End -->
       <%end if%>      
       
        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
		<!--Modify by Leo  新增判斷式  -- Start -->
			<% if (repeat) <> "1" then %>
            |<a href="<%=HTProgPrefix%>List.asp?epubid=<%=epubid%>&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>">下一頁</a> 
			<% else %>
			|<a href="<%=HTProgPrefix%>List.asp?epubid=<%=epubid%>&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>&repeat=1">下一頁</a> 
			<% end if %>
		<!--Modify by Leo  新增判斷式  -- End -->
        <%end if%>     
        | 每頁筆數:
       <select id=PerPage size="1" style="color:#FF0000">            
			 <!--<option value="1"<%if PerPageSize=1 then%> selected<%end if%>>1</option>-->
             <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
             <option value="100"<%if PerPageSize=100 then%> selected<%end if%>>100</option>                       
             <option value="200"<%if PerPageSize=200 then%> selected<%end if%>>200</option>
             <option value="300"<%if PerPageSize=300 then%> selected<%end if%>>300</option>
        </select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<A href="eSendListSearch.asp?epubid=<%=Request("epubid")%>" title="查詢電子報" align=right>查詢電子報</A>&nbsp;&nbsp;&nbsp;
		
		<!--Modify by Leo  新增判斷式  -- Start -->
		<% if (repeat) <> "1" then%>
		<A href="eSendListPrint.asp" title="匯出清單" align=right>匯出清單</A>&nbsp;&nbsp;&nbsp;
		<% else %>
		<A href="eSendListPrint.asp?repeat=1" title="匯出清單" align=right>匯出清單</A>&nbsp;&nbsp;&nbsp;
		<% end if %>
		<!--Modify by Leo  新增判斷式  -- End -->
		
		<!--Added by Leo  新增檢視 / 返回  -- Start -->
		<% if (repeat) <> "1" then %>
		<A href="eSendList.asp?epubid=<%=Request("epubid")%>&epTreeID=<%=Request("epTreeID")%>&repeat=1" title="檢視重複eMail" align="right">重複eMail</A>
		<% else %>
		<A href="eSendList.asp?epubid=<%=Request("epubid")%>&epTreeID=<%=Request("epTreeID")%>" title="返回發送清單" align="right">返回發送清單</A>
		<% end if %>
		<!--Added by Leo  新增檢視 / 返回  -- End -->
     </font>     
    <CENTER>
     <TABLE width=100% cellspacing="1" cellpadding="0" class=bg>                   
     <tr align=left>    
	<td class=eTableLable>eMail</td>
	<td class=eTableLable>發送日期</td>
	<td class=eTableLable>會員ID</td>
	<td class=eTableLable>姓名</td>
	<td class=eTableLable>自選訂閱</td>
    </tr>	                                
<%                   
    for i=1 to PerPageSize                  
%>                  
<tr>                  
<%pKey = ""
pKey = pKey & "&ePubID=" & RSreg("ePubID")
if pKey<>"" then  pKey = mid(pKey,2)%>
	<TD class=eTableContent><font size=2>
<%=RSreg("eMail")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("sendDate")%>
</font></td>
	<TD class=eTableContent><font size=2>
<A href="/member/member_edit.asp?account=<%=RSreg("account")%>">
<%=RSreg("account")%>
</A>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("realname")%>
</font></td>
	<TD class=eTableContent align=right><font size=2>
<% if RSreg("memEPCount")> 0 then	response.write RSreg("memEPCount")	%>
</font></td>
    </tr>
    <%
         RSreg.moveNext
         if RSreg.eof then exit for 
    next 
   %>
    </TABLE>
    </CENTER>
       <!-- 程式結束 ---------------------------------->  
     </form>  
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 
</form>
</body>
</html>                                 
<%else%>
      <script language=vbs>
           msgbox "找不到資料, 請重設查詢條件!"
	       window.history.back
      </script>
<%end if%>

<script Language=VBScript>
  dim gpKey
     sub GoPage_OnChange      
           newPage=reg.GoPage.value
		   <!--Modify by Leo  新增判斷式  -- Start -->
		   repeat = "<%=repeat%>"
		   if repeat = "" then
           document.location.href="<%=HTProgPrefix%>List.asp?epubid=<%=epubid%>&epTreeID=<%=Request("epTreeID")%>&nowPage=" & newPage & "&pagesize=<%=PerPageSize%>"
		   else 
		   document.location.href="<%=HTProgPrefix%>List.asp?epubid=<%=epubid%>&epTreeID=<%=Request("epTreeID")%>&nowPage=" & newPage & "&pagesize=<%=PerPageSize%>&repeat=1"
		   end if
		   <!--Modify by Leo  新增判斷式  -- End -->
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
		   <!--Modify by Leo  新增判斷式  -- Start -->
		   repeat = "<%=repeat%>"
		   if repeat = "" then
           document.location.href="<%=HTProgPrefix%>List.asp?epubid=<%=epubid%>&epTreeID=<%=Request("epTreeID")%>&nowPage=<%=nowPage%>&pagesize=" & newPerPage
		   else 
		   document.location.href="<%=HTProgPrefix%>List.asp?epubid=<%=epubid%>&epTreeID=<%=Request("epTreeID")%>&nowPage=<%=nowPage%>&pagesize=" & newPerPage & "&repeat=1"
		   end if
		   <!--Modify by Leo  新增判斷式  -- End -->
     end sub 

     sub setpKey(xv)
     	gpKey = xv
     end sub


</script>
