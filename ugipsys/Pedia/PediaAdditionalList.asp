<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="PEDIA01"
HTProgPrefix="PediaAdditional" 

Dim iCtUnitId : iCtUnitId = session("PediaAdditionalUnitId")
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<link href="/css/list.css" rel="stylesheet" type="text/css">
	<link href="/css/layout.css" rel="stylesheet" type="text/css">
	<title>資料管理／資料上稿</title>
</head>

<%
	if request("submitTask")="DELETE" then
		sql = ""
    For Each x In Request.Form
			if left(x,4)="dbox" and request(x) <> "" then
		    xn = mid(x, 5)
				sql = sql & "UPDATE Pedia SET xStatus = 'D' WHERE gicuitem = " & xn & ";"				
      end if
    Next 		
		conn.execute(sql)
		
		Dim flag : flag = false
		sql = "SELECT * FROM Activity WHERE ActivityId = 'pedia' AND GETDATE() BETWEEN ActivityStartTime AND ActivityEndTime"
		set rs = conn.execute(sql)
		if not rs.eof then
			flag = true
		end if
		
		if flag then
			For Each x In Request.Form
				if left(x,4)="dbox" and request(x) <> "" then
					xn = mid(x, 5)
					sql = "SELECT memberId FROM Pedia WHERE gicuitem = " & xn		
					set rs = conn.execute(sql)
					if not rs.eof then
						sql = "UPDATE ActivityPediaMember SET commendAdditionalCount = commendAdditionalCount - 1 WHERE memberId = '" & rs("memberId") & "' AND commendAdditionalCount > 0"
						conn.execute(sql)				
					end if
					rs.close
					set rs = nothing				
				end if
			Next 		
		end if		
		
%>
		<script language=VBS>
			alert "刪除完成！"
			window.navigate "<%=HTProgPrefix%>List.asp?icuitem=<%=request.querystring("icuitem")%>"
		</script>
<%	   
		response.end
	
	elseif request("submitTask")= "QUERY" then
	
		cSql = "SELECT COUNT(*) FROM Pedia INNER JOIN CuDtGeneric ON CuDtGeneric.icuitem = Pedia.gicuitem "
		cSql = cSql & " INNER JOIN Member ON Pedia.memberId = Member.account WHERE CuDtGeneric.ictunit = " & iCtUnitId 
		if request.querystring("icuitem") <> "" then
			cSql = cSql & " AND Pedia.parentIcuitem = " & request.querystring("icuitem")
		end if
		if request("xBody") <> "" then cSql = cSql & " AND xBody LIKE '%" & request("xBody") & "%' "
		if request("xPostDateS") <> "" then cSql = cSql & " AND commendTime >= '" & request("xPostDateS") & " 00:00:00' "
		if request("xPostDateE") <> "" then cSql = cSql & " AND commendTime <= '" & request("xPostDateE") & " 23:59:59' "
		if request("fCTUPublic") <> "" then cSql = cSql & " AND fCTUPublic = '" & request("fCTUPublic") & "' "
		if request("xStatus") <> "" then cSql = cSql & " AND xStatus = '" & request("xStatus") & "' "
		if request("memberId") <> "" then 
			cSql = cSql & " AND ( memberId LIKE '%" & request("memberId") & "%' OR realname LIKE '" & request("memberId") & "' "
			cSql = cSql & " OR nickname LIKE '%" & request("memberId") & "%') "
		end if		
		
		fSql = " CuDTGeneric.iCUItem, CuDTGeneric.sTitle, Pedia.memberId, Member.realname, Member.nickname, "
		fSql = fSql & " Pedia.xStatus, Pedia.parentIcuitem, Pedia.commendTime FROM CuDTGeneric INNER JOIN Pedia "
		fSql = fSql & " ON CuDTGeneric.iCUItem = Pedia.gicuitem INNER JOIN Member ON Pedia.memberId = Member.account "
		fSql = fSql & " WHERE CuDTGeneric.iCTUnit = " & iCtUnitId 
		if request.querystring("icuitem") <> "" then
			fSql = fSql & " AND Pedia.parentIcuitem = " & request.querystring("icuitem")
		end if
		if request("xBody") <> "" then fSql = fSql & " AND xBody LIKE '%" & request("xBody") & "%' "
		if request("xPostDateS") <> "" then fSql = fSql & " AND commendTime >= '" & request("xPostDateS") & " 00:00:00' "
		if request("xPostDateE") <> "" then fSql = fSql & " AND commendTime <= '" & request("xPostDateE") & " 23:59:59' "
		if request("fCTUPublic") <> "" then fSql = fSql & " AND fCTUPublic = '" & request("fCTUPublic") & "' "
		if request("xStatus") <> "" then fSql = fSql & " AND xStatus = '" & request("xStatus") & "' "
		if request("memberId") <> "" then 
			fSql = fSql & " AND ( memberId LIKE '%" & request("memberId") & "%' OR realname LIKE '" & request("memberId") & "' "
			fSql = fSql & " OR nickname LIKE '%" & request("memberId") & "%') "
		end if
		
		session("baseSql") = fSql
		session("cSql") = cSql
	
	elseif request("keep")="" then
   	
		cSql = "SELECT COUNT(*) FROM Pedia INNER JOIN CuDtGeneric ON CuDtGeneric.icuitem = Pedia.gicuitem "
		cSql = cSql & "WHERE CuDtGeneric.ictunit = " & iCtUnitId 
		if request.querystring("icuitem") <> "" then
			cSql = cSql & " AND Pedia.parentIcuitem = " & request.querystring("icuitem")
		end if
		'cSql = cSql & " ORDER BY CuDTGeneric.xPostDate DESC "
		
		fSql = " CuDTGeneric.iCUItem, CuDTGeneric.sTitle, Pedia.memberId, Member.realname, Member.nickname, "
		fSql = fSql & " Pedia.xStatus, Pedia.parentIcuitem, Pedia.commendTime FROM CuDTGeneric INNER JOIN Pedia "
		fSql = fSql & " ON CuDTGeneric.iCUItem = Pedia.gicuitem INNER JOIN Member ON Pedia.memberId = Member.account "
		fSql = fSql & " WHERE CuDTGeneric.iCTUnit = " & iCtUnitId 
		if request.querystring("icuitem") <> "" then
			fSql = fSql & " AND Pedia.parentIcuitem = " & request.querystring("icuitem")
		end if
		fSql = fSql & " ORDER BY CuDTGeneric.xPostDate DESC "
		
		session("baseSql") = fSql
		session("cSql") = cSql
	
	end if
	
	Dim sTitle : sTitle = ""
	if request.querystring("icuitem") <> "" then
		sql = "SELECT sTitle FROM CuDtGeneric WHERE icuitem = " & request.querystring("icuitem")
		Set rs = conn.execute(sql)
		if not rs.eof then
			sTitle = rs("sTitle")
		end if
		Set rs = nothing
	else
		sTitle = "全部補充解釋"
	end if
	
	fSql = session("baseSql")	
	cSql = session("cSql")
	
	nowPage = Request.QueryString("nowPage")  '現在頁數
  PerPageSize = cint(Request.QueryString("pagesize"))
  if PerPageSize <= 0 then  PerPageSize = 15  

  set RSc = conn.execute(cSql)
  totRec = RSc(0)       '總筆數
  totPage = int(totRec / PerPageSize + 0.999)

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
	<h1>資料管理／小百科補充管理</h1><font size=2>【目錄樹節點: 知識小百科】
	<div id="Nav">
		<A href="/Pedia/PediaQuery.asp?ictunit=<%=iCtUnitId%>&picuitem=<%=request.querystring("icuitem")%>" title="指定查詢條件">查詢</A>		
	  <a href="/Pedia/PediaList.asp">回條列頁</a>
		<% if request.querystring("icuitem") <> "" then %>
			<a href="/Pedia/PediaEdit.asp?icuitem=<%=request.querystring("icuitem")%>&phase=edit">回資料內容頁</a>	
		<% end if %>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">
	詞目標題：<%=sTitle%>&nbsp;
	<font size=2>【知識小百科補充】</font>
</div>
	
<Form id="Form1" name="reg" method="POST" action="<%=HTProgPrefix%>List.asp?icuitem=<%=request.querystring("icuitem")%>">
	<INPUT TYPE=hidden name=submitTask value="">	
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
		<th class="First" scope="col">&nbsp;</th>
		<th class="eTableLable">補充解釋</th>
		<th class="eTableLable">發表帳號</th>
		<th class="eTableLable">真實姓名</th>
		<th class="eTableLable">暱稱</th>
		<th class="eTableLable">狀態</th>
		<th class="eTableLable">發布日</th>		
	</tr>
<%
	If not RSreg.eof then   
		for i = 1 to PerPageSize
			pKey = "icuitem=" & RSreg("icuitem")
%>                  
			<tr>                  
				<td class=eTableContent><input type="checkbox" name="dbox<%=RSreg("icuitem")%>" value="Y"></td>
				<td class=eTableContent><A href="/Pedia/PediaAdditionalEdit.asp?<%=pKey%>&picuitem=<%=RSreg("parentIcuitem")%>&phase=edit&nowpage=<%=nowPage%>&pagesize=<%=PerPageSize%>"><%=RSreg("sTitle")%></A></td>
				<td class=eTableContent><%=RSreg("memberId")%></td>
				<td class=eTableContent><%=RSreg("realname")%></td>
				<td class=eTableContent><%=RSreg("nickname")%></td>
				<td class=eTableContent><%=RSreg("xStatus")%></td>
				<td class=eTableContent><%=RSreg("commendTime")%></td>
			</tr>
<%
      RSreg.moveNext
      if RSreg.eof then exit for 
		next 
%>
	</table>    
	<div align="center">
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
		
		dim gpKey
    sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=" & newPage & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>" & "&pagesize=" & newPerPage                    
     end sub 

     sub shortLongList(sorl)
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>" _
			& "&shortLongList=" & sorl
     end sub         
	</script>                           
<% else %>
  <script language=vbs>
		msgbox "找不到資料, 請重設查詢條件!"
		'window.history.back
  </script>
<% end if %>
