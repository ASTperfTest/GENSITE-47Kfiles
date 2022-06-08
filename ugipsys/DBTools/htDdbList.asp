<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="資料來源管理"
HTProgFunc="查詢"
HTProgCode="HT011"
HTProgPrefix="HtDdb" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 fSql=Request.QueryString("strSql")

if fSql="" then
	fSql = "SELECT b.* FROM HtDdb AS b WHERE 1=1"
	if request.form("htx_dbID") <> "" then
		whereCondition = replace("b.dbId LIKE N'%{0}%'", "{0}", request.form("htx_dbID") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_dbName") <> "" then
		whereCondition = replace("b.dbName LIKE N'%{0}%'", "{0}", request.form("htx_dbName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_dbDesc") <> "" then
		whereCondition = replace("b.dbDesc LIKE N'%{0}%'", "{0}", request.form("htx_dbDesc") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_dbIP") <> "" then
		whereCondition = replace("b.dbIp LIKE N'%{0}%'", "{0}", request.form("htx_dbIP") )
		fSql = fSql & " AND " & whereCondition
	end if
	for each x in request.form
	    if request(x) <> "" then
            if mid(x,2,3) = "fx_" then
	    	   select case left(x,1)
		          case "s"
		            	fSql = fSql & " AND " & mid(x,5) & "=" & pkStr(request(x),"")
		          case else
		            	fSql = fSql & " AND " & mid(x,5) & " LIKE N'%" & request(x) & "%'"
		         end select
	        end if
	    end if
	next 
end if

nowPage=Request.QueryString("nowPage")  '現在頁數
'response.write fSql
'response.end

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
      strSql=server.URLEncode(fSql)
   end if    
end if   
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="580" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>【查詢結果清單】</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right>
    		<!--Button OnClick="Javascript: window.open('../ws/WS_LocateStuQuery.asp','','width=600,height=350,scrollbars=yes');">學生Locate WS</Button-->
	       <%if (HTProgRight and 4)=4 then%>
	       <a href="<%=HTprogPrefix%>Add.asp">新增</a>	    
	       <%end if%>
	       <%if (HTProgRight and 2)=2 then%>
	       <a href="<%=HTprogPrefix%>Query.asp">重設查詢</a>	    
	       <%end if%>
    </td>    
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="80%" colspan="2" height=230 valign=top>
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
     <TABLE width=95% cellspacing="1" cellpadding="0" class=bg>                   
     <tr align=left>    
	<td align=center class=lightbluetable>資料庫代碼</td>
	<td align=center class=lightbluetable>資料庫名稱</td>
	<td align=center class=lightbluetable>說明</td>
	<td align=center class=lightbluetable>資料庫IP</td>
	<td align=center class=lightbluetable>Report</td>
    </tr>	                                
<%                   
    for i=1 to PerPageSize                  
%>                  
<tr>                  
<%pKey = ""
pKey = pKey & "&dbID=" & RSreg("dbID")
if pKey<>"" then  pKey = mid(pKey,2)%>
	<TD class=whitetablebg align=center><font size=2>
<%=RSreg("dbID")%>
</font></td>
	<TD class=whitetablebg align=center><font size=2>
	<A href="HtDdbEdit.asp?<%=pKey%>">
<%=RSreg("dbName")%>
</A>
</font></td>
	<TD class=whitetablebg align=center><font size=2>
<%=RSreg("dbDesc")%>
</font></td>
	<TD class=whitetablebg align=center><font size=2>
<%=RSreg("dbIP")%>
</font></td>
	<TD class=whitetablebg align=center><font size=2>
	<A href="ShowAFPList.asp?<%=pKey%>">關聯表</A>
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
