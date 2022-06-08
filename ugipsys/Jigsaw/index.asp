<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>
<!--#include virtual = "/inc/server.inc" -->

<!--#include virtual = "/inc/client.inc" -->


<%
  if request("clear") = "1" then
	session("jigsql") = ""
end if	
	'response.write request("sql")
	'response.write request("sql2")
	session("jigsql")=""
	'response.write session("jigsqlindex")
	'response.write session("jigsqlindex1")
	
	if session("jigsqlindex")="" then
	    'response.write "a"
		sql="SELECT [iCUItem],[fCTUPublic] ,[sTitle],[xImportant],[xPostDate],[xPostDateEnd] FROM [mGIPcoanew].[dbo].[CuDTGeneric] where [iCTUnit] =2199 and [iBaseDSD]=7 and rss='N'"
	    if request("sTitle")<>"" then 
		   sql=sql&" and ([sTitle] LIKE '%"&request("sTitle")&"%')"
		end if 
		
		if request("Status")<>"" then 
		   sql=sql&" and ([fCTUPublic] = '"&request("Status")&"')"
		end if 
		
		if request("value(startDate)")<>"" and request("value(endDate)")<>"" then 
		
		   'sql=sql&" and ([xPostDate] between '"&request("value(startDate)")&"' and '"&request("value(endDate)")&"') or ([xPostDateEnd]  between '"&request("value(startDate)")&"' and '"&request("value(endDate)")&"') "
		   sql=sql&" and ([xPostDate] >= '"&request("value(startDate)")&"' and [xPostDateEnd] <= '"&request("value(endDate)")&"' )"
		else
			if request("value(startDate)")<>""  then 
			sql=sql&" and ([xPostDate] >= '"&request("value(startDate)")&"')"
			else
				if  request("value(endDate)")<>"" then 
				sql=sql&" and ([xPostDateEnd] <='"&request("value(endDate)")&"') "
			    end if
			end if  
		end if 
	     sql=sql&"order by Created_Date desc, xImportant desc"
		set rs=conn.execute(sql)
	    session("jigsqlindex")=sql
	    'response.write "c"
	else
		'response.write "b"
		sql=session("jigsqlindex")
	    set rs=conn.execute(sql)
	end if
	


   	if session("jigsqlindex1")="" then
	  'response.write "c"
		sql2=" SELECT count(*) FROM  CuDTGeneric where [iBaseDSD]=7 and [iCTUnit]=2199 and rss='N'"
		if request("sTitle")<>"" then 
		  sql2=sql2&" and ([sTitle] LIKE '%"&request("sTitle")&"%')"
		end if 
		
		if request("Status")<>"" then 
		sql2=sql2&" and ([fCTUPublic] = '"&request("Status")&"')"
		end if 
		
		if request("value(startDate)")<>"" and request("value(endDate)")<>"" then 
		    sql2=sql2&" and ([xPostDate] >= '"&request("value(startDate)")&"' and [xPostDateEnd] <= '"&request("value(endDate)")&"' )"
		   'sql2=sql2&" and ([xPostDate] between '"&request("value(startDate)")&"' and '"&request("value(endDate)")&"') or ([xPostDateEnd]  between '"&request("value(startDate)")&"' and '"&request("value(endDate)")&"') "
		else
			if request("value(startDate)")<>""  then 
			sql2=sql2&" and ([xPostDate] > '"&request("value(startDate)")&"')"
			else
				if  request("value(endDate)")<>"" then 
				sql2=sql2&" and ([xPostDateEnd]<'"&request("value(endDate)")&"') "
			    end if
			end if 
		end if 
		set rs2=conn.execute(sql2)
	   session("jigsqlindex1")=sql2
	   'response.write "dd"
   else
		'response.write "d"
		sql2=session("jigsqlindex1")
        set rs2=conn.execute(sql2)
   end if
	
	
	if int(request("Pagesize"))=0 then
	 totalpage = 1
	 page = 1
    
	else
	
	pagecount = 10
    if request("Pagesize")<>"" then pagecount=request("Pagesize")
	
	
	page = 1
	if request("page") <> "" then	page = request("page")
    if int(rs2(0)) <= int(pagecount)  then  page = 1
	 
	totalpage = 1
	
		if rs2(0) > 0 then
			totalpage = rs2(0) \ pagecount
			'response.write(totalpage)	
			if rs2(0) mod pagecount <> 0 then	totalpage = totalpage + 1
		end if
    end if

	%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	<link rel="stylesheet" href="/inc/setstyle.css">
	<link type="text/css" rel="stylesheet" href="../css/list.css">
	<link type="text/css" rel="stylesheet" href="../css/layout.css">
	<link type="text/css" rel="stylesheet" href="../css/setstyle.css">
<title>知識拼圖管理</title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>

	<td width="50%" class="FormName" align="left">農業推薦單元知識拼圖管理&nbsp;
	<font size=2>【專區清單】</td>
    <td width="50%" class="FormLink" align="right">
	<A href="subjectPubAdd.asp" title="新增">新增專區</A>
	<A href="subject_query01.asp" title="查詢">專區查詢</A>
	<!--A href="Javascript:window.history.back();" title="回前頁">回前頁</A-->
	</td>
  </tr>
  <tr>
	<td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
  <tr>
<td width="95%" colspan="2" height=230 valign=top>
 <Form name=reg method="POST" action=../ePaperList.asp>
  <p align="center">  
     <font size="2" color="rgb(63,142,186)"> 第
     <font size="2" color="#FF0000"><%=page%>/<%= totalpage %></font>頁|                      
        共<font size="2" color="#FF0000"><%= rs2(0) %></font>
       <font size="2" color="rgb(63,142,186)">筆| 跳至第       

                  
       
      <select name="select" id="select" size="1" style="color:#FF0000" onchange="search1()">
<%
	
	n = 1
	while n <= totalpage
		response.write "<option value='" & n & "'"
		if int(n) = int(page)  then	 response.write " selected" 
		response.write ">" & n & "</option>"
		n = n + 1
      
		
	wend
   	
%>	 </font> 頁 </select> 
	    | 每頁筆數:
		
       <select id="PageSize" name="PageSize" size="1" style="color:#FF0000" onchange="PageSizeOnChange()">       
                <option value="0" <% If pagecount = 0 Then %> selected <% End If %> >全部</option>                       
             	<option value="15" <% If pagecount = 15 Then %> selected <% End If %> >15</option>
             	<option value="30" <% If pagecount = 30 Then %> selected <% End If %> >30</option>
             	<option value="50" <% If pagecount = 50 Then %> selected <% End If %> >50</option>
        </select> 
<CENTER>
<table cellspacing="0" id="ListTable">
        <tr>
          <th>主題專區名稱</th>
          <th>日期區間</th>
          <th>重要性</th>
          <th>是否公開</th>
          <th>設定專區內容</th>
        </tr>
        
   <%

  
  i = 0	  
   while not rs.eof
   i = i + 1
		
		if  int(request("Pagesize"))=0 then
%>		
		 <tr> <td class=eTableContent><a href="index_subject.asp?iCUItem=<%=rs("iCUItem")%>"><%=rs("sTitle")%></a></td>
		 <%
		  xPostDate=right(year(rs("xPostDate")),4)   &"/"&   right("0"&month(rs("xPostDate")),2)&   "/"   &   right("0"&day(rs("xPostDate")),2)   
		  xPostDateEnd=right(year(rs("xPostDateEnd")),4)   &"/"&   right("0"&month(rs("xPostDateEnd")),2)&   "/"   &   right("0"&day(rs("xPostDateEnd")),2)   
		  %>
		  <td class=eTableContent><%=xPostDate%>-<%=xPostDateEnd%></td>
          <td align="center" class=eTableContent><%=rs("xImportant")%></td>
          <td align="center" class=eTableContent><%if rs("fCTUPublic")="Y" then%><%response.write "公開"%><%else  response.write "不公開"  end if %></td>
          <td align="center" class=eTableContent><font size=2> <a href="../ePubList.htm">
            <input name="button" type="button" class="cbutton" onClick="location.href='subjectPubList.asp?iCUItem=<%=rs("iCUItem")%>'" value="設定">
          </a> </font></td></tr>
		
		
<%		
		else
		if i > (page - 1) * pagecount and i <= page * pagecount Then

   %>
         
		 <tr> <td class=eTableContent><a href="index_subject.asp?iCUItem=<%=rs("iCUItem")%>"><%=rs("sTitle")%></a></td>
          <%
		  xPostDate=right(year(rs("xPostDate")),4)   &"/"&   right("0"&month(rs("xPostDate")),2)&   "/"   &   right("0"&day(rs("xPostDate")),2)   
		  xPostDateEnd=right(year(rs("xPostDateEnd")),4)   &"/"&   right("0"&month(rs("xPostDateEnd")),2)&   "/"   &   right("0"&day(rs("xPostDateEnd")),2)   
		  %>
		  <td class=eTableContent><%=xPostDate%>-<%=xPostDateEnd%></td>
          <td align="center" class=eTableContent><%=rs("xImportant")%></td>
          <td align="center" class=eTableContent><%if rs("fCTUPublic")="Y" then%><%response.write "公開"%><%else  response.write "不公開"  end if %></td>
          <td align="center" class=eTableContent><font size=2> <a href="../ePubList.htm">
            <input name="button" type="button" class="cbutton" onClick="location.href='subjectPubList.asp?iCUItem=<%=rs("iCUItem")%>'" value="設定">
          </a> </font></td></tr>
         <%
           end if
		   end if 
		rs.movenext
	        wend
			%>
   
	   
        
</table>
</CENTER>
       <!-- 程式結束 ---------------------------------->  
     </form>
	 </td>
  </tr>  
</table>
</body>
</html>                                 
<script language="JavaScript">
function search1(){
	var doc = document.reg;
	var objSelect =document.reg.select;
window.location.href='index.asp?page='+objSelect.options[objSelect.selectedIndex].value+'&PageSize='+ doc.PageSize.options[doc.PageSize.selectedIndex].value;
}
function PageSizeOnChange() 
{
	var doc = document.reg;
	var objSelect =document.reg.select;
    window.location.href='index.asp?page='+objSelect.options[objSelect.selectedIndex].value+'&PageSize='+ doc.PageSize.options[doc.PageSize.selectedIndex].value;
   
}
</script>

