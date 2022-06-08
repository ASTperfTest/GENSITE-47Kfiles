<%@ CodePage = 65001 %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap = "單元資料維護"
HTProgFunc = "資料清單"
HTProgCode = "AVote"
HTProgPrefix = "AttitudeVote" 

Dim iCtUnitId : iCtUnitId = "2180"

%>
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title></title>
</head>

<% Sub showDoneBox(lMsg, atype) %>
  <html>
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<meta name="GENERATOR" content="Hometown Code Generator 1.0">
			<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
			<title>編修表單</title>
    </head>
    <body>
			<script language=vbs>
			  alert("<%=lMsg%>")
				<% if atype = "1" then %>
					history.back()
				<% else %>
					document.location.href="<%=HTprogPrefix%>List.asp?nowpage=<%=request.querystring("nowpage")%>&pagesize=<%=request.querystring("pagesize")%>"
				<% end if %>
			</script>
    </body>
  </html>
<% End sub %>
<%
	if request("submitTask") = "RECOVER" then
	
		if request("ckbox") = "" then
			showDoneBox "請選擇要恢復的意見", "1"
			response.end
		end if
		Dim ckbox1 : ckbox1 = request("ckbox") & ","
		Dim items1 : items1 = split(ckbox1, ",")
		Dim count1 : count1 = 0
		for each item1 in items1
			if item1 <> "" then
				sql = "UPDATE CuDTGeneric SET fCTUPublic = 'Y' WHERE iCUItem = " & item1
				conn.execute(sql)
				count1 = count1 + 1
			end if
		next
		
		showDoneBox "恢復成功, 共恢復了 " & count1 & " 筆", "2"
		response.end
		
	elseif request("submitTask") = "DELETE" then
	
		if request("ckbox") = "" then
			showDoneBox "請選擇要刪除的意見", "1"
			response.end
		end if
		Dim ckbox : ckbox = request("ckbox") & ","
		Dim items : items = split(ckbox, ",")
		Dim count : count = 0
		for each item in items
			if item <> "" then
				sql = "UPDATE CuDTGeneric SET fCTUPublic = 'N' WHERE iCUItem = " & item
				conn.execute(sql)
				count = count + 1
			end if
		next
		
		showDoneBox "刪除成功, 共刪除了 " & count & " 筆", "2"
		response.end
		
	elseif request("submitTask")= "Query" then
	
		Dim stitle : stitle = request("sTitle")
		Dim xbody : xbody = request("xBody")
		Dim ieditor : ieditor = request("iEditor")
	
		cSql = "SELECT COUNT(*) FROM CuDtGeneric INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
		cSql = cSql & "WHERE CuDtGeneric.ictunit = " & iCtUnitId & " AND (CAST(xBody AS nvarchar) <> '') "
		if stitle <> "" then cSql = cSql & " AND sTitle LIKE '%" & stitle & "%' "
		if xbody <> "" then cSql = cSql & " AND CAST(xBody AS nvarchar) LIKE '%" & xbody & "%' "
		if ieditor <> "" then cSql = cSql & " AND iEditor LIKE '%" & ieditor & "%' "
			
		fSql = "CuDTGeneric.*, Member.account, Member.realname, Member.nickname "
		fSql = fSql & "FROM CuDTGeneric INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
		fSql = fSql & "WHERE CuDTGeneric.iCTUnit = " & iCtUnitId & " AND (CAST(xBody AS nvarchar) <> '') "
		if stitle <> "" then fSql = fSql & " AND sTitle LIKE '%" & stitle & "%' "
		if xbody <> "" then fSql = fSql & " AND CAST(xBody AS nvarchar) LIKE '%" & xbody & "%' "
		if ieditor <> "" then fSql = fSql & " AND iEditor LIKE '%" & ieditor & "%' "
		
		fSql = fSql & " ORDER BY CuDTGeneric.xPostDate DESC "
		
		session("baseSql") = fSql
		session("cSql") = cSql
				
  elseif request("keep") = "" then
   	
		cSql = "SELECT COUNT(*) FROM CuDTGeneric INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
		cSql = cSql & "WHERE CuDtGeneric.ictunit = " & iCtUnitId & " AND (CAST(xBody AS nvarchar) <> '')"
		
		fSql = "CuDTGeneric.*, Member.account, Member.realname, Member.nickname "
		fSql = fSql & "FROM CuDTGeneric INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
		fSql = fSql & "WHERE CuDTGeneric.iCTUnit = " & iCtUnitId & " AND (CAST(xBody AS nvarchar) <> '') "
		fSql = fSql & "ORDER BY CuDTGeneric.Created_Date DESC "
		
		session("baseSql") = fSql
		session("cSql") = cSql
	
	end if

	fSql = session("baseSql")	
	cSql = session("cSql")
	
	nowPage = Request.QueryString("nowPage")  '現在頁數
  PerPageSize = Request.QueryString("pagesize")
	if not isNumeric(PerPageSize) then
		PerPageSize = 10
	else
		PerPageSize = cint(Request.QueryString("pagesize"))
	end if
  if PerPageSize <= 0 then  PerPageSize = 10  

  set RSc = conn.execute(cSql)
  totRec = RSc(0)       '總筆數
  totPage = int(totRec / PerPageSize + 0.999)

	if not isNumeric(nowPage) then
		nowPage = 1
	else
		nowPage = cint(Request.QueryString("nowPage"))
	end if
	
  if cint(nowPage) < 1 then 
    nowPage = 1
  elseif cint(nowPage) > totPage then 
    nowPage = totPage 
  end if            	

	fSql = "SELECT TOP " & nowPage * PerPageSize & fSql

	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

'----------HyWeb GIP DB CONNECTION PATCH----------
'	RSreg.Open fSql, Conn, 3, 1
Set RSreg =  Conn.execute(fSql)
	RSreg.CacheSize = PerPageSize
'----------HyWeb GIP DB CONNECTION PATCH----------


	if Not RSreg.eof then
		if totRec > 0 then 
      RSreg.PageSize = PerPageSize       '每頁筆數
      RSreg.AbsolutePage = nowPage      
		end if    
	end if   
%>

<body>
<div id="FuncName">
	<h1>資料管理／態度投票</h1><font size=2></font>
	<div id="Nav">
		<A href="AttitudeVoteQuery.asp" title="指定查詢條件">查詢</A>				
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">	
	<font size=2>【主題單元:態度投票】</font>
</div>
<Form id="Form2" name=reg method="POST" action="<%=HTprogPrefix%>List.asp?nowpage=<%=request.querystring("nowpage")%>&pagesize=<%=request.querystring("pagesize")%>" >
	<INPUT TYPE=hidden name="submitTask" value="">	
	<div id="Page">    
    <% if cint(nowPage) <> 1 then %>
			<img src="/images/arrow_previous.gif" alt="上一頁">       		
			<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>">上一頁</a> ，
    <% end if %>      
		共<em><%=totRec%></em>筆資料，每頁顯示
    <select id=PerPage size="1" style="color:#FF0000" class="select">            
			<option value="15"<%if PerPageSize=15 then%> selected<%end if%>>15</option>                       
      <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
      <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
      <option value="300"<%if PerPageSize=300 then%> selected<%end if%>>300</option>
    </select>     
		筆，目前在第
    <select id=GoPage size="1" style="color:#FF0000" class="select">
		<% For iPage=1 to totPage %> 
			<option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
    <% Next %>   
    </select>      
		頁	
    <% if cint(nowPage)<>totPage then %> 
     ，<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>">下一頁
      <img src="/images/arrow_next.gif" alt="下一頁"></a> 
    <% end if %>     
	</div>

	<table cellspacing="0" id="ListTable">
	<tr>		
		<th><p align="center">&nbsp;</p></th>
		<th><p align="center">文章標題</p></th>
		<th><p align="center">投票意見</th>		
		<th><p align="center">投票日期</th>
		<th><p align="center">會員名稱</th>
		<th><p align="center">父節點ID</th>
		<th><p align="center">狀態</th>
	</tr>
<%
	If not RSreg.eof then   
		for i = 1 to PerPageSize						
			pKey = "icuitem=" & RSreg("icuitem")
%>                  
			<tr>                  				
				<td class=eTableContent><input type="checkbox" name="ckbox" value="<%=RSreg("icuitem")%>"></td>					
				<td class=eTableContent><%=RSreg("sTitle")%></td>
				<td class=eTableContent><%=RSreg("xBody")%></td>
				<td class=eTableContent><%=RSreg("xPostDate")%></td>				
				<td class=eTableContent>
				<%
						response.write RSreg("account") & "&nbsp; | &nbsp;" & RSreg("realname") & "&nbsp; | &nbsp;" & RSreg("nickname")
				%>
				</td>
				<td class=eTableContent><%=RSreg("refID")%></td>				
				<td class=eTableContent>
				<%
						if RSreg("fCTUPublic") = "Y" then
							response.write "開放"
						else
							response.write "刪除"
						end if
				%>
				</td>				
			</tr>
<%
      RSreg.moveNext
      if RSreg.eof then exit for 
		next 
%>
	</table>     
	<div align="center">
		<input name="button1" type=button class="cbutton" onClick="formModSubmit1()" value ="整批恢復">	
		<input name="button" type=button class="cbutton" onClick="formModSubmit()" value ="整批刪除">	
		<!--input name="button" type=button class="cbutton" onClick="javascript:window.location.href='/Pedia/PediaEdit.asp?icuitem=<%=request.querystring("icuitem")%>&phase=edit'" value ="回前頁"-->	
	</div>  
  </form>  
</body>
</html>      
<script Language=VBScript>
	
	sub formModSubmit()		
		reg.submitTask.value = "DELETE"
		reg.action = "<%=HTProgPrefix%>List.asp?icuitem=<%=request.querystring("icuitem")%>"
		reg.Submit
	end sub

	sub formModSubmit1()		
		reg.submitTask.value = "RECOVER"
		reg.action = "<%=HTProgPrefix%>List.asp?icuitem=<%=request.querystring("icuitem")%>"
		reg.Submit
	end sub
	
  sub GoPage_OnChange      
    newPage=reg.GoPage.value     
    document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=" & newPage & "&pagesize=<%=PerPageSize%>"    
  end sub      
     
  sub PerPage_OnChange                
    newPerPage=reg.PerPage.value
    document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>" & "&pagesize=" & newPerPage                    
	end sub 
</script>                                       
<% else %>
  <script language=vbs>
		msgbox "找不到資料, 請重設查詢條件!"
		'window.history.back
  </script>
<% end if %>
