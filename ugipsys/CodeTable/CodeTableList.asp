<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="代碼定義管理"
HTProgFunc="代碼定義清單"
HTProgCode="Pn50M03"
HTProgPrefix="CodeTable" %>
<!--#INCLUDE FILE="CodeTableListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 'fSql=Request.QueryString("strSql")
  If Request.QueryString = "" Then
    session("strSql") = ""
  End If
  fSql = session("strSql")

if fSql="" then
	fSql = "SELECT htx.codeId, htx.codeName, htx.codeTblName, htx.codeType, htx.codeRank, htx.showOrNot, xref1.mvalue AS xref1showOrNot" _
		& " FROM (CodeMetaDef AS htx LEFT JOIN CodeMain AS xref1 ON xref1.mcode = showOrNot AND xref1.codeMetaId=N'boolYN')" _
		& " WHERE 1=1"
	xpCondition
	fSql = fSql & " ORDER BY " & "htx.codeType, htx.codeRank, htx.codeId"
end if

nowPage=Request.QueryString("nowPage")  '現在頁數


'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=20  
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
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right>
		<%if (HTProgRight and 2)=2 then%>
			<A href="CodeTableQuery.asp" title="指定查詢條件">查詢</A>
		<%end if%>
		<%if (HTProgRight and 4)=4 then%>
			<A href="CodeTableAdd.asp" title="新增代碼定義">新增</A>
		<%end if%>
    </td>    
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
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
	<td class=eTableLable>代碼類別</td>
	<td class=eTableLable>代碼ID</td>
	<td class=eTableLable>代碼名稱</td>
	<td class=eTableLable>順序</td>
	<td class=eTableLable>存放資料表</td>
	<td class=eTableLable>是否顯示</td>
    </tr>	                                
<%                   
    for i=1 to PerPageSize                  
%>                  
<tr>                  
<%pKey = ""
pKey = pKey & "&codeId=" & RSreg("codeId")
if pKey<>"" then  pKey = mid(pKey,2)%>
	<TD class=eTableContent><font size=2>
<%=RSreg("codeType")%>
</font></td>
	<TD class=eTableContent><font size=2>
	<A href="CodeTableEdit.asp?<%=pKey%>">
<%=RSreg("codeId")%>
</A>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("codeName")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("codeRank")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("codeTblName")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("xref1showOrNot")%>
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
