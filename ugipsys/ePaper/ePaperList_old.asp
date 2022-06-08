<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="電子報管理"
HTProgFunc="電子報清單"
HTProgCode="GW1M51"
HTProgPrefix="ePaper"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 fSql=Request.QueryString("strSql")
'response.Write fSql &"**1st<HR>"
if fSql="" then
	fSql = "SELECT htx.ctRootId,htx.ctRootName, htx.purpose " _
		& " ,(Select count(*) from EpPub EP where EP.ctRootId=htx.ctRootId) EpPubCount " _
		& " FROM CatTreeRoot AS htx " _
		& " WHERE htx.vgroup='GE'"		
end if
'response.Write fSql &"**2nd<HR>"
'response.End
nowPage=Request.QueryString("nowPage")  '現在頁數


'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,3,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------

'RSreg.Open fSql,Conn
'totRec=RSreg.Recordcount
'response.Write "totRec:" & totRec & Conn

if Not RSreg.eof then 
	'response.Write "111"
   totRec=RSreg.Recordcount       '總筆數
   'response.Write totRec
   'PerPageSize = 10
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

      'RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
   end if
   'response.Write " Not RSreg.eof" & RSreg("ctRootId")
   'response.Write PerPageSize
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
	<td class=eTableLable>電子報名稱</td>
	<td class=eTableLable>用途</td>
	<td class=eTableLable>發行期數</td>
	<td class=eTableLable>發行清單</td>
    </tr>	                                
<%      
	           
    for i=1 to PerPageSize                  
    'response.Write RSreg("ctRootId")  
%>                  
<tr>                  
	<TD class=eTableContent><font size=2>
	<A href="ePaperEdit.asp?epTreeID=<%=RSreg("ctRootId")%>&ePaperName=<%=server.urlencode(RSreg("ctRootName"))%>">
<%=RSreg("ctRootName")%>
</A>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("purpose")%>
</font></td>
	<TD class=eTableContent><font size=2><%=RSreg("EpPubCount")%></font></td>
	<TD class=eTableContent><font size=2>
	<A href="ePubList.asp?epTreeID=<%=RSreg("ctRootId")%>&ePaperName=<%=server.urlencode(RSreg("ctRootName"))%>">檢視清單</A>
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
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>&pagesize=" & newPerPage                    
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
