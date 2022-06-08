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
	Dim filename : filename = replace( DateAdd("d", -1 , year(now) & "/" & month(now) & "/" & day(now) ), "/" , "")
	Response.AddHeader "Content-Disposition", "attachment;filename=" & filename & ".xls" 
	Response.ContentType = "application/vnd.ms-excel" 

	Dim account : account = request.querystring("account")
	Dim realname : realname = request.querystring("realname")
	Dim nickname : nickname = request.querystring("nickname")
	Dim homeaddr : homeaddr = request.querystring("homeaddr")
	Dim mobile : mobile = request.querystring("mobile")
	Dim email : email = request.querystring("email")
	Dim id : id = request.querystring("id")
	Dim phone : phone = request.querystring("phone")
	Dim level : level = request.querystring("level")
	Dim gradetype : gradetype = request.querystring("gradetype")
	Dim gradefrom : gradefrom = request.querystring("gradefrom")
	Dim gradeto : gradeto = request.querystring("gradeto")
	Dim condition : condition = ""
	Dim gradetypethis : gradetypethis = request.querystring("gradetypethis")
	
	sqlSelect = "SELECT MemberGradeSummary.memberId, MemberGradeSummary.browseTotal, MemberGradeSummary.loginTotal, " & _
							"MemberGradeSummary.shareTotal, MemberGradeSummary.contentTotal, MemberGradeSummary.additionalTotal, " & _
							"MemberGradeSummary.calculateTotal, Member.realname, ISNULL(Member.nickname, '') AS nickname,Member.homeaddr, Member.mobile, Member.email, Member.phone, Member.id "
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
	if id <> "" then 
		condition = condition & "身分證字號：" & id & ","
		sqlWhere = sqlWhere & "AND Member.id LIKE '%" & id & "%' "
	end if
	if homeaddr <> "" then 
		condition = condition & "住址：" & homeaddr & ","
		sqlWhere = sqlWhere & "AND Member.homeaddr LIKE '%" & homeaddr & "%' "
	end if
	
	if phone <> "" then 
		condition = condition & "電話(住家)：" & phone & ","
		sqlWhere = sqlWhere & "AND Member.phone LIKE '%" & phone & "%' "
	end if
	if mobile <> "" then 
		condition = condition & "電話(手機)：" & mobile & ","
		sqlWhere = sqlWhere & "AND Member.mobile LIKE '%" & mobile & "%' "
	end if
	if email <> "" then 
		condition = condition & "Email：" & email & ","
		sqlWhere = sqlWhere & "AND Member.email LIKE '%" & email & "%' "
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
	
	if gradetypethis = "1" then
		sqlOrder = "ORDER BY MemberGradeSummary.calculateTotal DESC"
	elseif gradetypethis = "2" then
		sqlOrder = "ORDER BY MemberGradeSummary.browseTotal DESC, MemberGradeSummary.calculateTotal DESC"
	elseif gradetypethis = "3" then
		sqlOrder = "ORDER BY MemberGradeSummary.loginTotal DESC, MemberGradeSummary.calculateTotal DESC"
	elseif gradetypethis = "4" then
		sqlOrder = "ORDER BY MemberGradeSummary.shareTotal DESC, MemberGradeSummary.calculateTotal DESC"
	elseif gradetypethis = "5" then
		sqlOrder = "ORDER BY MemberGradeSummary.contentTotal DESC, MemberGradeSummary.calculateTotal DESC"
	elseif gradetypethis = "6" then
		sqlOrder = "ORDER BY MemberGradeSummary.additionalTotal DESC, MemberGradeSummary.calculateTotal DESC"
	else 
		sqlOrder = "ORDER BY MemberGradeSummary.calculateTotal DESC"
	end if
	
	if len(condition) > 0 then condition = left(condition, len(condition) - 1)
	
	response.write "<table border=""1"">" & vbcrlf
	response.write "<tr><td colspan=""14""><font face=""新細明體"">匯出日期：" & Date() & "</font></td></tr>" & vbcrlf 
	response.write "<tr><td colspan=""14""><font face=""新細明體"">條件 => " & condition & "</font></td></tr>" & vbcrlf
	response.write "<tr><td colspan=""14""><font face=""新細明體"">&nbsp;</font></td>"
	response.write "<tr><td><font face=""新細明體"">帳號</font></td><td><font face=""新細明體"">真實姓名</font></td>" & _
								 "<td><font face=""新細明體"">暱稱</font></td><td><font face=""新細明體"">身分證字號</font></td><td><font face=""新細明體"">Email</font></td><td><font face=""新細明體"">電話(住家)</font></td><td><font face=""新細明體"">電話(手機)</font></td><td><font face=""新細明體"">住址</font></td>" & _
	"<td><font face=""新細明體"">總得分</font></td>" & _
	"<td><font face=""新細明體"">瀏覽得點</font></td><td><font face=""新細明體"">登入得點</font></td>" & _
								 "<td><font face=""新細明體"">互動得點</font></td><td><font face=""新細明體"">內容價值</font></td>" & _
								 "<td><font face=""新細明體"">參與活動</font></td></tr>" & vbcrlf
	
	set rs = conn.execute(sqlSelect & sqlFrom & sqlWhere & sqlOrder)	
	while not rs.eof
		response.write "<tr><td><font face=""新細明體"">&nbsp;" & rs("memberId") & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & trim(rs("realname")) & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & trim(rs("nickname")) & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & trim(rs("id")) & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & trim(rs("email")) & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & trim(rs("phone")) & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & trim(rs("mobile")) & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & trim(rs("homeaddr")) & "</font></td>" & _
									 "<td><font face=""新細明體"">" & rs("calculateTotal") & "</font></td>" & _
									 "<td><font face=""新細明體"">" & rs("browseTotal") & "</font></td>" & _
									 "<td><font face=""新細明體"">" & rs("loginTotal") & "</font></td>" & _
									 "<td><font face=""新細明體"">" & rs("shareTotal") & "</font></td>" & _
									 "<td><font face=""新細明體"">" & rs("contentTotal") & "</font></td>" & _
									 "<td><font face=""新細明體"">" & rs("additionalTotal") & "</font></td></tr>" & vbcrlf
		rs.movenext
	wend
	rs.close
	set rs = nothing
	
	response.write "</table>" & vbcrlf
	
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
%>
