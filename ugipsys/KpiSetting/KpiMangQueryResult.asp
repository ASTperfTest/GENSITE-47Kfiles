<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	HTProgCap = "單元資料維護"
	HTProgFunc = "查詢"
	HTProgCode="kpi01"
	HTProgPrefix="" 
	response.charset = "UTF-8"
%>
<!--#include virtual = "/inc/server.inc" -->
<%
	
	Dim account 
	Dim realname 
	Dim nickname 
	Dim level 
	Dim gradetype 
	Dim gradefrom 
	Dim gradeto 
	Dim condition 
	Dim gradetypethis : gradetypethis = request.querystring("gradetypethis")
	Dim exporturl
	
	Function GetGradeByLevel( level )
		Dim grade : grade = 0
		if level > 4 then 
			grade = 99999
		else
			sql = "SELECT TOP 1 * FROM CodeMain WHERE codeMetaID = 'gradeLevel' AND mCode = '" & level & "'"
			set rs = conn.execute(sql)
			if not rs.eof then
				grade = rs("mValue")
			end if
			rs.close
			set rs = nothing
		end if
		GetGradeByLevel = grade
	End Function
	
	Function GetNameByGrade( grade )
		Dim name : name = ""
		Dim level : level = ""
		if grade = "0" then
			level = "1"
		else
			sql = "SELECT TOP 1 * FROM CodeMain WHERE codeMetaID = 'gradeLevel' AND " & grade & " < mValue ORDER BY mSortValue"
			set rs = conn.execute(sql)
			if not rs.eof then
				level = rs("mCode")
			else
				level = "5"
			end if
			rs.close
			set rs = nothing
			level = CStr(CInt(level) - 1)
		end if
		if level = "1" then
			name = "入門級"
		elseif level = "2" then
			name = "進階級"
		elseif level = "3" then
			name = "高手級"
		elseif level = "4" then
			name = "達人級"
		end if
		GetNameByGrade = name
	End Function	
	
	if request.querystring("keep") = "" then
	
		account = request.querystring("account")
		realname = request.querystring("realname")
		nickname = request.querystring("nickname")
		level = request.querystring("level")
		gradetype = request.querystring("gradetype")
		gradefrom = request.querystring("gradefrom")
		gradeto = request.querystring("gradeto")	
			
		sqlSelect = "MemberGradeSummary.memberId, MemberGradeSummary.browseTotal, MemberGradeSummary.loginTotal, " & _
								"MemberGradeSummary.shareTotal, MemberGradeSummary.contentTotal, MemberGradeSummary.additionalTotal, " & _
								"MemberGradeSummary.calculateTotal, Member.realname, ISNULL(Member.nickname, '') AS nickname "
		sqlFrom = "FROM MemberGradeSummary INNER JOIN Member ON MemberGradeSummary.memberId = Member.account "
		sqlWhere = "WHERE 1 = 1 "
				
		if account <> "" then 
			condition = condition & "帳號：" & account & ","
			sqlWhere = sqlWhere & "AND Member.account LIKE '%" & account & "%' "
		end if
		if realname <> "" then 
			condition = condition & "真實姓名：" & realname & ","
			sqlWhere = sqlWhere & "AND Member.realname LIKE '%" & realname & "%' "
		end if
		if nickname <> "" then 
			condition = condition & "暱稱：" & nickname & ","
			sqlWhere = sqlWhere & "AND Member.nickname LIKE '%" & nickname & "%' "
		end if
		if level <> "" then 
			if level = "1" then
				name = "入門級"
			elseif level = "2" then
				name = "進階級"
			elseif level = "3" then
				name = "高手級"
			elseif level = "4" then
				name = "達人級"
			end if
			condition = condition & "會員等級：" & name & ","
			sqlWhere = sqlWhere & "AND MemberGradeSummary.calculateTotal >= " & GetGradeByLevel(CInt(level)) & " "
			sqlWhere = sqlWhere & "AND MemberGradeSummary.calculateTotal < " & GetGradeByLevel(CInt(level) + 1) & " "
		end if
		if gradefrom <> "" or gradeto <> "" then 
			if gradetype = "1" then 
				condition = condition & "分數類別：會員總積分,"
				sqlWhere = sqlWhere & "AND MemberGradeSummary.calculateTotal "
			elseif gradetype = "2" then 
				condition = condition & "分數類別：吸收度(瀏覽行為)積分,"
				sqlWhere = sqlWhere & "AND MemberGradeSummary.browseTotal "
			elseif gradetype = "3" then 
				condition = condition & "分數類別：持久度(登入行為)積分,"
				sqlWhere = sqlWhere & "AND MemberGradeSummary.loginTotal "
			elseif gradetype = "4" then 
				condition = condition & "分數類別：分享度(互動行為)積分,"
				sqlWhere = sqlWhere & "AND MemberGradeSummary.shareTotal "
			elseif gradetype = "5" then 
				condition = condition & "分數類別：貢獻度(內容價值)積分,"
				sqlWhere = sqlWhere & "AND MemberGradeSummary.contentTotal "
			elseif gradetype = "6" then 
				condition = condition & "分數類別：踴躍度(活動參與)積分,"
				sqlWhere = sqlWhere & "AND MemberGradeSummary.additionalTotal "
			end if
			if gradefrom <> "" and gradeto <> "" then
				condition = condition & "會員總積分：" & gradefrom & "~" & gradeto & ","
				sqlWhere = sqlWhere & "BETWEEN " & gradefrom & " AND " & gradeto & " "
			elseif gradefrom <> "" and gradeto = "" then
				condition = condition & "會員總積分：大於 " & gradefrom & ","
				sqlWhere = sqlWhere & ">= " & gradefrom & " "
			elseif gradefrom = "" and gradeto <> "" then
				condition = condition & "會員總積分：小於 " & gradeto & ","
				sqlWhere = sqlWhere & "<= " & gradeto & " "
			end if						
		end if
		
		if len(condition) > 0 then condition = left(condition, len(condition) - 1)
	
		session("csql") = "SELECT COUNT(*) " & sqlFrom & sqlWhere
		session("fsql") = sqlSelect & sqlFrom & sqlWhere 
		session("condition") = condition
		exporturl = "KpiMemberSummaryExport.asp?account=" & Server.URLEncode(account) & "&realname=" & Server.URLEncode(realname) & _
								"&nickname=" & Server.URLEncode(nickname) & "&level=" & level & "&gradetype=" & gradetype & _
								"&gradefrom=" & gradefrom & "&gradeto=" & gradeto
		session("exporturl") = exporturl
	end if
	if gradetypethis = "1" then
		sqlOrder = "ORDER BY MemberGradeSummary.calculateTotal DESC"
		exporturl = session("exporturl") & "&gradetypethis=1"
	elseif gradetypethis = "2" then
		sqlOrder = "ORDER BY MemberGradeSummary.browseTotal DESC, MemberGradeSummary.calculateTotal DESC"
		exporturl = session("exporturl") & "&gradetypethis=2"
	elseif gradetypethis = "3" then
		sqlOrder = "ORDER BY MemberGradeSummary.loginTotal DESC, MemberGradeSummary.calculateTotal DESC"
		exporturl = session("exporturl") & "&gradetypethis=3"
	elseif gradetypethis = "4" then
		sqlOrder = "ORDER BY MemberGradeSummary.shareTotal DESC, MemberGradeSummary.calculateTotal DESC"
		exporturl = session("exporturl") & "&gradetypethis=4"
	elseif gradetypethis = "5" then
		sqlOrder = "ORDER BY MemberGradeSummary.contentTotal DESC, MemberGradeSummary.calculateTotal DESC"
		exporturl = session("exporturl") & "&gradetypethis=5"
	elseif gradetypethis = "6" then
		sqlOrder = "ORDER BY MemberGradeSummary.additionalTotal DESC, MemberGradeSummary.calculateTotal DESC"
		exporturl = session("exporturl") & "&gradetypethis=6"
	else 
		sqlOrder = "ORDER BY MemberGradeSummary.calculateTotal DESC"
		exporturl = session("exporturl") & "&gradetypethis=1"
	end if
	fSql = session("fsql") & sqlOrder	
	cSql = session("cSql")
	condition = session("condition")

	nowPage = Request.QueryString("nowPage")  '現在頁數
  PerPageSize = Request.QueryString("pagesize")
	if not isNumeric(PerPageSize) then
		PerPageSize = 15
	else
		PerPageSize = cint(Request.QueryString("pagesize"))
	end if
  if PerPageSize <= 0 then PerPageSize = 15

  set RSc = conn.execute(cSql)
  totRec = RSc(0)       '總筆數
  totPage = int(totRec / PerPageSize + 0.999)

  if cint(nowPage) < 1 then 
    nowPage = 1
  elseif cint(nowPage) > totPage then 
    nowPage = totPage 
  end if            	

	fSql = "SELECT TOP " & nowPage * PerPageSize & " " & fSql 

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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<link href="/css/form.css" rel="stylesheet" type="text/css" /></head>
<title>資料管理／資料上稿</title>
</head>
<body>
	<div id="FuncName">
		<h1>總年度KPI整合管理</h1>
		<div id="Nav">
			<A href="KpiMangQuery.asp" title="重新查詢">重新查詢</A>	
		</div>
		<div id="ClearFloat"></div>
	</div>
	<div id="FormName">會員KPI積分管理</div>

	<Form id="Form2" name="reg" method="POST" action="">
	<INPUT TYPE="hidden" name="submitTask" value="" />	
	<div class="browseby">
		查詢條件【<%=condition%>】| <a href="<%=exporturl%>" target="_blank">產生此查詢結果積分總表</a><br />
		瀏覽排序依： 
		<SELECT name="gradetypethis" size="1" onchange="gradetypeonchange()">
			<option value="1" <% if gradetypethis = "" then %> selected <% end if %>>會員總積分</option>
			<option value="2" <% if gradetypethis = "2" then %> selected <% end if %>>吸收度(瀏覽行為)積分</option>
			<option value="3" <% if gradetypethis = "3" then %> selected <% end if %>>持久度(登入行為)積分</option>
			<option value="4" <% if gradetypethis = "4" then %> selected <% end if %>>分享度(互動行為)積分</option>
			<option value="5" <% if gradetypethis = "5" then %> selected <% end if %>>貢獻度(內容價值)積分</option>
			<option value="6" <% if gradetypethis = "6" then %> selected <% end if %>>踴躍度(活動參與)積分</option>
		</SELECT>
	</div>
	<div id="Page">    
    <% if cint(nowPage) <> 1 then %>
			<img src="/images/arrow_previous.gif" alt="上一頁">       		
			<a href="KpiMangQueryResult.asp?keep=Y&gradetypethis=<%=gradetypethis%>&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>">上一頁</a> ，
    <% end if %>      
		共<em><%=totRec%></em>筆資料，每頁顯示
    <select id="PerPage" size="1" style="color:#FF0000" class="select" onchange="perpageonchange()">            
			<option value="15"<%if PerPageSize=15 then%> selected<%end if%>>15</option>                       
      <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
      <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
      <option value="300"<%if PerPageSize=300 then%> selected<%end if%>>300</option>
    </select>     
		筆，目前在第
    <select id="GoPage" size="1" style="color:#FF0000" class="select" onchange="gopageonchange()">
		<% For iPage=1 to totPage %> 
			<option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
    <% Next %>   
    </select>      
		頁	
    <% if cint(nowPage)<>totPage then %> 
     ，<a href="KpiMangQueryResult.asp?keep=Y&gradetypethis=<%=gradetypethis%>&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>">下一頁
      <img src="/images/arrow_next.gif" alt="下一頁"></a> 
    <% end if %>     
	</div>
	<table cellspacing="0" id="ListTable">
	<tr>
		<th scope="col">會員帳號</th>
		<th scope="col">真實姓名</th>
		<th scope="col">暱稱</th>
		<th scope="col">會員等級</th>
		<th scope="col" width="10%">總積分</th>
		<th scope="col" width="10%">吸收度(瀏覽)</th>
		<th scope="col" width="10%">持久度(登入)</th>
		<th scope="col" width="10%">分享度(互動)</th>
		<th scope="col" width="10%">貢獻度(內容)</th>
		<th scope="col" width="10%">踴躍度(活動)</th>
	</tr>   
	<%
	If not RSreg.eof then   
		for i = 1 to PerPageSize
			response.write "<tr>" & vbcrlf
			response.write "<td><a href=""/member/newMemberEdit.asp?account=" & trim(RSreg("memberId")) & """>" & trim(RSreg("memberId")) & "&nbsp;</a></td>" & vbcrlf 
			response.write "<td>" & trim(RSreg("realname")) & "&nbsp;</td>" & vbcrlf 
			response.write "<td>" & trim(RSreg("nickname")) & "&nbsp;</td>" & vbcrlf 
			response.write "<td>" & GetNameByGrade(RSreg("calculateTotal")) & "&nbsp;</td>" & vbcrlf 
			response.write "<td><a href=""KpiCalculateTotal.asp?memberid=" & trim(RSreg("memberId")) & """>" & RSreg("calculateTotal") & "&nbsp;</a></td>" & vbcrlf 
			response.write "<td><a href=""KpiBrowseTotal.asp?memberid=" & trim(RSreg("memberId")) & """>" & RSreg("browseTotal") & "&nbsp;</a></td>" & vbcrlf 
			response.write "<td><a href=""KpiLoginTotal.asp?memberid=" & trim(RSreg("memberId")) & """>" & RSreg("loginTotal") & "&nbsp;</a></td>" & vbcrlf 
			response.write "<td><a href=""KpiShareTotal.asp?memberid=" & trim(RSreg("memberId")) & """>" & RSreg("shareTotal") & "&nbsp;</a></td>" & vbcrlf 
			response.write "<td>" & RSreg("contentTotal") & "&nbsp;</td>" & vbcrlf 
			response.write "<td>" & RSreg("additionalTotal") & "&nbsp;</td>" & vbcrlf 
			response.write "</tr>" & vbcrlf 
			RSreg.moveNext
      if RSreg.eof then exit for 
		next 
	end if
	%>
	</table>
	</form>  
</body>
</html> 
<script language="javascript">
	function gradetypeonchange()
	{		
		window.location.href = "KpiMangQueryResult.asp?keep=Y&gradetypethis=" + document.getElementById("gradetypethis").value + "&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>";
	}
	function gopageonchange()
	{
		window.location.href = "KpiMangQueryResult.asp?keep=Y&gradetypethis=<%=gradetypethis%>&nowPage=" + document.getElementById("GoPage").value + "&pagesize=<%=PerPageSize%>";    
  }
	function perpageonchange()
  {                
		window.location.href = "KpiMangQueryResult.asp?keep=Y&gradetypethis=<%=gradetypethis%>&nowPage=<%=nowPage%>&pagesize=" + document.getElementById("PerPage").value;
  }

</script>














