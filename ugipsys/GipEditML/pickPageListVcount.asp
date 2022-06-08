<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="選取文章"
HTProgFunc="符合清單"
HTProgCode="GC1AP1"
HTProgPrefix="pickPage" %>
<!--#INCLUDE FILE="pickPageListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%

if request("keep")="" then
'	fSql = "SELECT htx.CtUnitID, htx.CtUnitName, htx.CtUnitKind, htx.iBaseDSD, htx.fCtUnitOnly, xref1.mValue AS xref1CtUnitKind, xref2.sBaseDSDName AS xref2iBaseDSD, xref3.mValue AS xref3fCtUnitOnly, d.deptName" _
'		& " FROM (((CtUnit AS htx LEFT JOIN CodeMainLong AS xref1 ON xref1.mCode = htx.CtUnitKind AND xref1.codeMetaID='refCTUKind') 
'		LEFT JOIN BaseDSD AS xref2 ON xref2.iBaseDSD = htx.iBaseDSD AND xref2.inUse='Y') LEFT JOIN CodeMain AS xref3 ON xref3.mCode = htx.fCtUnitOnly AND xref3.codeMetaID='boolYN')" _
'		& " LEFT JOIN Dept AS d ON d.deptID=htx.deptID " _
'		& " WHERE 1=1 " 
	xSelect = " htx.iCUItem, htx.sTitle, htx.xImgFile, u.CtUnitName"
	fSql = " FROM CuDtGeneric As htx JOIN CtUnit AS u ON htx.iCtUnit=u.CtUnitID" _
		& "  JOIN CatTreeNode As n ON htx.iCtUnit=n.CtUnitID" _
		& " WHERE 1=1 " 
'	fSql = fSql & " AND (htx.deptID IS NULL OR htx.deptID LIKE '" & session("deptID") & "%')" 
	xpCondition
	session("baseSql") = fSql
	session("xSelectSql") = xSelect
end if
	fSql = session("baseSql")
	xSelect = session("xSelectSql")
	cSql = "SELECT count(*) " & fSql

nowPage=Request.QueryString("nowPage")  '現在頁數
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  PerPageSize=15  
      
      set RSc = conn.execute(cSql)
	  totRec=RSc(0)       '總筆數
      response.write totRec & "<HR>"
      totPage = int(totRec/PerPageSize+0.999)
      response.write totPage & "<HR>"
      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) >totPage then 
         nowPage=totPage 
      end if            	

	fSql = fSql & " ORDER BY CtUnitName, sTitle"
	fSql = "SELECT DISTINCT TOP " & nowPage*PerPageSize & xSelect & fSql
'	response.write fSql
'	response.end


 Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,3,1
Set RSreg = Conn.execute(fSql)
	RSreg.CacheSize = PerPageSize
'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      
      RSreg.PageSize=PerPageSize       '每頁筆數
      RSreg.AbsolutePage=nowPage
      strSql=server.URLEncode(fSql)
   end if    
end if   
%>
<body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
<% if (HTProgRight and 2)=2 then %><a href="pickPageQuery.asp">查詢</a>　<% End If %>
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
			<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>">上一頁</a> ，
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
            ，<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>">下一頁
            <img src="/images/arrow_next.gif" alt="下一頁"></a> 
        <%end if%>     
	</div>

	<!-- 列表 -->
	<table cellspacing="0" id="ListTable">
     <tr>    
	<th>主題單元</th>
	<th>標題</th>
	<th>單位</th>
    </tr>	                                
<%                   
    for i=1 to PerPageSize                  
%>                  
<tr>                  
<%pKey = ""
pKey = pKey & "&iCUItem=" & RSreg("iCUItem")
if pKey<>"" then  pKey = mid(pKey,2)%>
	<TD class="Center"><font size=2>
<%=RSreg("CtUnitName")%>
</font></td>
	<TD class=eTableContent><font size=2>
	<A href="CtUnitEdit.asp?<%=pKey%>"><%=RSreg("sTitle")%></A>
</font></td>
	<TD class="Center"><font size=2>
<%=RSreg("xImgFile")%>
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
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=" & newPage & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>" & "&pagesize=" & newPerPage                    
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
