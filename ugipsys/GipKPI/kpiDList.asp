<%@ CodePage = 65001 %>
<% Response.Expires = 0
response.charset = "utf-8"
HTProgCap="績效管理"
HTProgFunc="查詢"
HTProgCode="GC1AP9"
HTProgPrefix="kpiD" %>
<!--#INCLUDE FILE="kpiDListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 'fSql=Request.QueryString("strSql")
  If Request.QueryString("nowPage") = "" Then
    session("strSql") = ""
  End If
  fSql = session("strSql")

if fSql="" then
	fSql = "SELECT htx.*, Dept.abbrName, xref1.mvalue AS xtopCat, xref2.mvalue AS xPublic, cu.ctUnitName, u.userName" _
		& " FROM CuDtGeneric AS htx" _
		& " LEFT JOIN CodeMain AS xref1 ON xref1.mcode = htx.topCat AND xref1.codeMetaId='topDataCat'" _
		& " LEFT JOIN CodeMain AS xref2 ON xref2.mcode = htx.fctupublic AND xref2.codeMetaId='isPublic3'" _
		& " LEFT JOIN CtUnit AS cu ON cu.ctUnitId = htx.ictunit" _
		& " LEFT JOIN InfoUser AS u ON u.userId = htx.ieditor" _
		& " LEFT JOIN Dept  ON Dept.deptId = htx.idept " _
		& " WHERE 1=1 " 

	xConStr = ""
	xpCondition
	session("xConStr") = xConStr
	fSql = fSql & " ORDER BY htx.topCat, htx.idept, htx.deditDate DESC" 
end if

nowPage=Request.QueryString("nowPage")  '現在頁數
'response.write fSql

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
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
      'strSql=server.URLEncode(fSql)
	  session("strSql") = fSql
   end if    
end if   
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
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
		<%if (HTProgRight and 1)=1 then%>
			<A href="kpiQuery.asp" title="重設查詢條件">查詢</A>
		<%end if%>
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A> 
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext"><table border=0><%=session("xConStr")%></table>
	    </td>    
	    <td class="Formtext" align=right></td>
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
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">上一頁</a> 
       <%end if%>      
       
        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">下一頁</a> 
        <%end if%>     
        | 每頁筆數:
       <select id=PerPage size="1" style="color:#FF0000">            
             <option value="10"<%if PerPageSize=10 then%> selected<%end if%>>10</option>                       
             <option value="20"<%if PerPageSize=20 then%> selected<%end if%>>20</option>
             <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
             <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
        </select>     
     </font>     
    <CENTER>
     <TABLE width=100% cellspacing="1" cellpadding="0" class=bg>                   
     <tr align=left>    
	<td class=eTableLable>主題單元</td>
	<td class=eTableLable>標　題</td>
	<td class=eTableLable>狀 態</td>
	<td class=eTableLable>編修者</td>
	<td class=eTableLable>編修日</td>
    </tr>	                                
<%                   
    for i=1 to PerPageSize        
    	xUrl = session("myWWWSiteURL") & "/content.asp?CuItem="  & RSreg("iCuItem")
%>                  
<tr>                  
	<TD class=eTableContent><font size=2>
<%=RSreg("ctUnitName")%>
</font></td>
	<TD class=eTableContent><font size=2>
	<A href="<%=xUrl%>" target="_wMof">
<%=RSreg("sTitle")%>
</A>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("xPublic")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("userName")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=d7date(RSreg("deditDate"))%>
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
'	       window.history.back
      </script>
<%end if%>

<script Language=VBScript>
  dim gpKey
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&strSql=" & strSql & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>" & "&strSql=<%=strSql%>&pagesize=" & newPerPage                    
     end sub 

     sub setpKey(xv)
     	gpKey = xv
     end sub

	sub butAction(k) 
	  if gpKey = "" then 
	    alert("請先選擇個案!") 
	  else
    	select case k 
    	end select 
	  end if 
	end sub 

</script>
