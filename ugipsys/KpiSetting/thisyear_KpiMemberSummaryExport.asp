<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	HTProgCap = "單元資料維護"
	HTProgFunc = "查詢"
	HTProgCode="kpi01"
	HTProgPrefix="" 
	response.charset = "UTF-8"
	Response.Buffer = False
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
	
	' sqlSelect = "SELECT MemberGradeSummary.memberId, MemberGradeSummary.browseTotal, MemberGradeSummary.loginTotal, " & _
							' "MemberGradeSummary.shareTotal, MemberGradeSummary.contentTotal, MemberGradeSummary.additionalTotal, " & _
							' "MemberGradeSummary.calculateTotal, Member.realname, ISNULL(Member.nickname, '') AS nickname,Member.homeaddr, Member.mobile, Member.email, Member.phone, Member.id "
	sqlSelect = "SELECT isnull(GradeBrowse_thisyear.GradeBrowse,0) as GradeBrowse, isnull(GradeLogin_thisyear.GradeLogin,0) as GradeLogin, "
	sqlSelect = sqlSelect & " isnull(GradeShare_thisyear.GradeShare,0) as GradeShare, Member.account as memberId , (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)+ISNULL(MGCBY.contentIQPoint,0)+ISNULL(MGCBY.contentRaisePoint,0)+ISNULL(MGCBY.contentLogoPoint,0)+ ISNULL(MGCBY.contentPlantLogPoint,0)+ISNULL(MGCBY.contentFishBowlPoint,0)+ ISNULL(MGCBY.contentDrPoint,0)+ ISNULL(MGCBY.contentKnowledgeQAPoint,0)+ ISNULL(MGCBY.contentDr2010Point,0)+ISNULL(MGCBY.contentTreasureHuntPoint,0)) as contentTotal, isnull(VIEWADDTHIS.ThisYearADDpoint, 0) as additionalTotal ,"
	sqlSelect = sqlSelect & " (isnull(GradeBrowse_thisyear.GradeBrowse,0) * 0.15 + isnull(GradeLogin_thisyear.GradeLogin,0) * 0.2) + isnull(GradeShare_thisyear.GradeShare,0) * 0.3 + (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)+ISNULL(MGCBY.contentIQPoint,0)+ISNULL(MGCBY.contentRaisePoint,0)+ISNULL(MGCBY.contentLogoPoint,0)+ ISNULL(MGCBY.contentPlantLogPoint,0)+ISNULL(MGCBY.contentFishBowlPoint,0)+ ISNULL(MGCBY.contentDrPoint,0)+ ISNULL(MGCBY.contentKnowledgeQAPoint,0)+ ISNULL(MGCBY.contentDr2010Point,0)+ISNULL(MGCBY.contentTreasureHuntPoint,0)) * 0.20 + isnull(VIEWADDTHIS.ThisYearADDpoint, 0) * 0.15 AS total , "
	sqlSelect = sqlSelect & " Member.realname, ISNULL(Member.nickname, '') AS nickname ,Member.homeaddr, Member.mobile, Member.email, Member.phone, Member.id "
	
	sqlFrom = " FROM Member "
	sqlFrom = sqlFrom & " LEFT JOIN GradeBrowse_thisyear ON Member.account = GradeBrowse_thisyear.memberId "
	sqlFrom = sqlFrom & " LEFT JOIN GradeLogin_thisyear ON Member.account = GradeLogin_thisyear.memberId "
	sqlFrom = sqlFrom & " LEFT JOIN GradeShare_thisyear ON Member.account = GradeShare_thisyear.memberId "
	sqlFrom = sqlFrom & " left JOIN MemberGradeSummary B ON Member.account = B.memberId "
	sqlFrom = sqlFrom & " LEFT JOIN GradeContentBrowse_thisyear VCB ON Member.account = VCB.ieditor  "
	sqlFrom = sqlFrom & " LEFT JOIN GradeContentCommend_thisyear VCC ON Member.account = VCC.ieditor  "
	sqlFrom = sqlFrom & " LEFT JOIN GradeContentDiscuss_thisyear VCD ON Member.account = VCD.ieditor "
	sqlFrom = sqlFrom & " LEFT JOIN View_MemberGradeAdditional_ThisYear VIEWADDTHIS ON Member.account = VIEWADDTHIS.memberId "
	sqlFrom = sqlFrom & " LEFT JOIN View_MemberGradeContent_ThisYear VIEWConTHIS ON Member.account = VIEWConTHIS.memberId "
	sqlFrom = sqlFrom & " LEFT JOIN dbo.MemberGradeContentByYear MGCBY ON Member.account = MGCBY.memberId AND (CONVERT(INT, CONVERT(varchar(4), getdate(), 120 ) ) - ISNULL(years,0) = 0) "
'	sqlFrom = "FROM MemberGradeSummary INNER JOIN Member ON MemberGradeSummary.memberId = Member.account "
	sqlWhere = " WHERE 1 = 1 "
	
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

		sqlWhere = sqlWhere & "AND (((isnull(GradeBrowse_thisyear.GradeBrowse,0) * 0.2) + (isnull(GradeLogin_thisyear.GradeLogin,0) * 0.2)) + (isnull(GradeShare_thisyear.GradeShare,0) * 0.3)) + (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)) * 0.15 + isnull(B.additionalTotal-B.additionalTotalHistory, 0) * 0.15 >= " & GetGradeByLevel(CInt(level)) & " "
		sqlWhere = sqlWhere & "AND (((isnull(GradeBrowse_thisyear.GradeBrowse,0) * 0.2) + (isnull(GradeLogin_thisyear.GradeLogin,0) * 0.2)) + (isnull(GradeShare_thisyear.GradeShare,0) * 0.3)) + (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)) * 0.15 + isnull(B.additionalTotal-B.additionalTotalHistory, 0) * 0.15 < " & GetGradeByLevel(CInt(level) + 1) & " "
	end if
	if gradefrom <> "" or gradeto <> "" then 
		if gradetype = "1" then 
			condition = condition & "分數類別：會員總積分,"
			sqlWhere = sqlWhere & "AND ((isnull(GradeBrowse_thisyear.GradeBrowse, 0) * 0.2 + isnull(GradeLogin_thisyear.GradeLogin, 0) * 0.2) + isnull(GradeShare_thisyear.GradeShare, 0) * 0.3 + (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)) * 0.15 + isnull(B.additionalTotal-B.additionalTotalHistory, 0) * 0.15) "
		elseif gradetype = "2" then 
			condition = condition & "分數類別：吸收度(瀏覽行為)積分,"
			sqlWhere = sqlWhere & "AND GradeBrowse_thisyear.GradeBrowse "
		elseif gradetype = "3" then 
			condition = condition & "分數類別：持久度(登入行為)積分,"
			sqlWhere = sqlWhere & "AND GradeLogin_thisyear.GradeLogin "
		elseif gradetype = "4" then 
			condition = condition & "分數類別：分享度(互動行為)積分,"
			sqlWhere = sqlWhere & "AND GradeShare_thisyear.GradeShare "
		 elseif gradetype = "5" then 
			 condition = condition & "分數類別：貢獻度(內容價值)積分,"
			 sqlWhere = sqlWhere & "AND (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)) "
		 elseif gradetype = "6" then 
			 condition = condition & "分數類別：踴躍度(活動參與)積分,"
			 sqlWhere = sqlWhere & "AND (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)) "
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
		sqlOrder = "ORDER BY Total DESC"
	elseif gradetypethis = "2" then
		sqlOrder = "ORDER BY GradeBrowse_thisyear.GradeBrowse DESC, Total DESC"
	elseif gradetypethis = "3" then
		sqlOrder = "ORDER BY GradeLogin_thisyear.GradeLogin DESC, Total DESC"
	elseif gradetypethis = "4" then
		sqlOrder = "ORDER BY GradeShare_thisyear.GradeShare DESC, Total DESC"
	elseif gradetypethis = "5" then
		sqlOrder = "ORDER BY contentTotal, Total DESC"
	elseif gradetypethis = "6" then
		sqlOrder = "ORDER BY additionalTotal, Total DESC"
	else 
		sqlOrder = "ORDER BY Total DESC"
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
									 "<td><font face=""新細明體"">" & CInt(rs("Total")) & "</font></td>" & _
									 "<td><font face=""新細明體"">" & rs("GradeBrowse") & "</font></td>" & _
									 "<td><font face=""新細明體"">" & rs("GradeLogin") & "</font></td>" & _
									 "<td><font face=""新細明體"">" & rs("GradeShare") & "</font></td>" & _
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
