<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題單元管理"
HTProgFunc="單元清單"
HTProgCode="GE1T11"
HTProgPrefix="CtUnit" %>
<!--#INCLUDE FILE="CtUnitListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 'fSql=Request.QueryString("strSql")
 If Request.QueryString = "" Then
    session("strSql") = ""
  End If
  fSql = session("strSql")

if fSql="" then
	fSql = "SELECT htx.redirectUrl, htx.ctUnitId, htx.ctUnitName, htx.ctUnitKind, htx.ibaseDsd, htx.fctUnitOnly, xref1.mvalue AS xref1ctUnitKind, xref2.sbaseDsdname AS xref2ibaseDsd, xref3.mvalue AS xref3fctUnitOnly, d.deptName" _
		& " ,(SELECT count(*) FROM CatTreeNode WHERE ctUnitId=htx.ctUnitId) AS NodeCount " _
		& " ,(SELECT count(*) FROM CuDtGeneric WHERE ictunit=htx.ctUnitId) AS ItemCount " _
		& " FROM (((CtUnit AS htx LEFT JOIN CodeMainLong AS xref1 ON xref1.mcode = htx.ctUnitKind AND xref1.codeMetaId='refCTUKind') LEFT JOIN BaseDsd AS xref2 ON xref2.ibaseDsd = htx.ibaseDsd AND xref2.inUse='Y') LEFT JOIN CodeMain AS xref3 ON xref3.mcode = htx.fctUnitOnly AND xref3.codeMetaId='boolYN')" _
		& " LEFT JOIN Dept AS d ON d.deptId=htx.deptId " _
		& " WHERE 1=1 " 
	fSql = fSql & " AND (htx.deptId IS NULL OR htx.deptId LIKE '" & session("deptId") & "%')" 
	xpCondition
	fSql = fSql & " ORDER BY htx.ctUnitName, ctUnitKind, htx.ibaseDsd"
'	response.write fsql
'	response.end
end if
nowPage=Request.QueryString("nowPage")  '現在頁數


'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,3,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=30  
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
<body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
<% if (HTProgRight and 4)=4 then %><a href="CtUnitAdd.asp?phase=add">新增</a>　<% End If %>
<% if (HTProgRight and 2)=2 then %><a href="CtUnitQuery.asp">查詢</a>　<% End If %>
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A> 
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName"><%=HTProgFunc%></div>
 <Form id="Form2" name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
	<!-- 分頁 -->
	<div id="Page">
       <% if cint(nowPage) <>1 then %>
			<img src="/images/arrow_previous.gif" alt="上一頁">       		
			<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">上一頁</a> ，
       <%end if%>      
		共<em><%=totRec%></em>筆資料，每頁顯示
       <select id=PerPage size="1" style="color:#FF0000" class="select">            
             <option value="15"<%if PerPageSize=15 then%> selected<%end if%>>15</option>                       
             <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
             <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
             <option value="300"<%if PerPageSize=300 then%> selected<%end if%>>300</option>
        </select>     
		筆，目前在第
         <select id=GoPage size="1" style="color:#FF0000" class="select">
             <%For iPage=1 to totPage%> 
                   <option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
             <%Next%>   
         </select>      
		頁	
        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            ，<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">下一頁
            <img src="/images/arrow_next.gif" alt="下一頁"></a> 
        <%end if%>     
	</div>

	<!-- 列表 -->
	<table cellspacing="0" id="ListTable">
     <tr>    
	<th>主題單元名稱</th>
	<th>單元類型</th>
	<th>單元資料定義/URL連結</th>
	<th>單位</th>
	<th>項目節點數</th>
	<th>資料筆數</th>
    </tr>	                                
<%                   
    for i=1 to PerPageSize                  
%>                  
<tr>                  
<%pKey = ""
pKey = pKey & "&ctUnitId=" & RSreg("ctUnitId")
if pKey<>"" then  pKey = mid(pKey,2)%>
	<TD class=eTableContent><font size=2>
	<A href="CtUnitEdit.asp?<%=pKey%>&phase=edit">
<%=RSreg("ctUnitName")%>
</A>
</font></td>
	<TD class="Center"><font size=2>
<%=RSreg("xref1ctUnitKind")%>
</font></td>
	<TD class="Center"><font size=2>
<% if RSreg("xref1ctUnitKind") = "URL 連結" then %>
	<%=RSreg("redirectUrl")%>
<% else %>
	<%=RSreg("xref2ibaseDsd")%>
<% end if %>
</font></td>
	<TD  class="Center"><%=RSreg("deptName")%></td>
	<TD class="Center"><font size=2>
	<A href="CtUnitNodeList.asp?<%=pKey%>" target="_blank">
<%=RSreg("NodeCount")%></A>
</font></td>
	<TD class="Center"><font size=2>
<%=RSreg("ItemCount")%>
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
