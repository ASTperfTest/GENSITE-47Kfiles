<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	HTProgCap = "單元資料維護"
	HTProgFunc = "查詢"
	HTUploadPath = session("Public") + "data/"
	HTProgCode="kpi01"
	HTProgPrefix="" 
	response.charset = "UTF-8"
%>
<!--#include virtual = "/inc/server.inc" -->
<%
	Dim memberid : memberid = request.querystring("memberid")
	Dim memberstr : memberstr = ""
	Dim thisyear : thisyear = request.querystring("thisyear")
	sql = "SELECT account, realname, nickname FROM Member WHERE account = '" & memberid & "'"
	set rs = conn.execute(sql)
	if not rs.eof then
		memberstr = trim(rs("account")) & "｜" & trim(rs("realname")) & "｜" & trim(rs("nickname"))
	end if 
	rs.close
	set rs = nothing	
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	<link rel="stylesheet" href="/inc/setstyle.css">
	<link type="text/css" rel="stylesheet" href="/css/list.css">
	<link type="text/css" rel="stylesheet" href="/css/layout.css">
	<link type="text/css" rel="stylesheet" href="/css/setstyle.css">
<title></title>
</head>

<body>
	<div id="FuncName">
		<%if thisyear <> "" then%>
		<h1>本年度KPI整合管理</h1>
		<%else%>
		<h1>總年度KPI整合管理</h1>
		<%end if%>
		<div id="Nav">
			<%if thisyear <> "" then%>
			<A href="thisyear_KpiMangQuery.asp" title="重新查詢">重設本年度KPI查詢</A>
			<%else%>
			<A href="KpiMangQuery.asp" title="重新查詢">重設總年度KPI查詢</A>
			<%end if%>	
			<A href="javascript:history.go(-1)" title="回前頁">回前頁</A>	
		</div>
		<div id="ClearFloat"></div>
	</div>
	<div id="FormName">會員KPI積分內容</div>

	<Form id="Form2" name="reg" method="POST" action="">
	<INPUT TYPE="hidden" name="submitTask" value="" />	
	<div class="browseby">會員：【<%=memberstr%>】</div>
	<table cellspacing="0" id="ListTable">
  <tr>
    <th width="16%" scope="col">總積分</th>
    <th width="16%" scope="col">吸收度(瀏覽)</th>
    <th width="16%" scope="col">持久度(登入)</th>
    <th width="16%" scope="col">分享度(互動)</th>
    <th width="16%" scope="col">貢獻度(內容)</th>
    <th scope="col">踴躍度(活動)</th>
  </tr>
	<%
	Dim calculateTotal, browseTotal, loginTotal, shareTotal, contentTotal, additionalTotal
	if thisyear = "" then
	sql = "SELECT * FROM MemberGradeSummary WHERE memberId = '" & memberid & "'"
	elseif thisyear <> "" then
	sql = "SELECT isnull(GradeBrowse_thisyear.GradeBrowse,0) as GradeBrowse, isnull(GradeLogin_thisyear.GradeLogin,0) as GradeLogin, "
	sql = sql & " isnull(GradeShare_thisyear.GradeShare,0) as GradeShare, Member.account , (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)+ISNULL(MGCBY.contentIQPoint,0)+ISNULL(MGCBY.contentRaisePoint,0)+ISNULL(MGCBY.contentLogoPoint,0)+ ISNULL(MGCBY.contentPlantLogPoint,0)+ISNULL(MGCBY.contentFishBowlPoint,0)+ ISNULL(MGCBY.contentDrPoint,0)+ ISNULL(MGCBY.contentKnowledgeQAPoint,0)+ ISNULL(MGCBY.contentDr2010Point,0)+ISNULL(MGCBY.contentTreasureHuntPoint,0)) as contentTotal, isnull(VIEWADDTHIS.ThisYearADDpoint, 0) as additionalTotal ,"
	sql = sql & " (isnull(GradeBrowse_thisyear.GradeBrowse,0) * 0.15 + isnull(GradeLogin_thisyear.GradeLogin,0) * 0.2) + isnull(GradeShare_thisyear.GradeShare,0) * 0.3 + (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)+ISNULL(MGCBY.contentIQPoint,0)+ISNULL(MGCBY.contentRaisePoint,0)+ISNULL(MGCBY.contentLogoPoint,0)+ ISNULL(MGCBY.contentPlantLogPoint,0)+ISNULL(MGCBY.contentFishBowlPoint,0)+ ISNULL(MGCBY.contentDrPoint,0)+ ISNULL(MGCBY.contentKnowledgeQAPoint,0)+ ISNULL(MGCBY.contentDr2010Point,0)+ISNULL(MGCBY.contentTreasureHuntPoint,0)) * 0.20 + isnull(VIEWADDTHIS.ThisYearADDpoint, 0) * 0.15 AS total "
	sql = sql & " FROM Member "
	sql = sql & " LEFT JOIN GradeBrowse_thisyear ON Member.account = GradeBrowse_thisyear.memberId "
	sql = sql & " LEFT JOIN GradeLogin_thisyear ON Member.account = GradeLogin_thisyear.memberId "
	sql = sql & " LEFT JOIN GradeShare_thisyear ON Member.account = GradeShare_thisyear.memberId "
	sql = sql & " left JOIN MemberGradeSummary B ON Member.account = B.memberId  "
	sql = sql & " LEFT JOIN GradeContentBrowse_thisyear VCB ON Member.account = VCB.ieditor  "
	sql = sql & " LEFT JOIN GradeContentCommend_thisyear VCC ON Member.account = VCC.ieditor  "
	sql = sql & " LEFT JOIN GradeContentDiscuss_thisyear VCD ON Member.account = VCD.ieditor "
	sql = sql & " LEFT JOIN View_MemberGradeAdditional_ThisYear VIEWADDTHIS ON Member.account = VIEWADDTHIS.memberId "
	sql = sql & " LEFT JOIN dbo.MemberGradeContentByYear MGCBY ON Member.account = MGCBY.memberId AND (CONVERT(INT, CONVERT(varchar(4), getdate(), 120 ) ) - ISNULL(years,0) = 0) "
	sql = sql & " where Member.account ='"& memberid &"'"
	end if
	set rs = conn.execute(sql)
	if not rs.eof then
		if thisyear <> "" then
			calculateTotal = CInt(rs("Total"))
			browseTotal = rs("GradeBrowse")
			loginTotal = rs("GradeLogin")
			shareTotal = rs("GradeShare")
			contentTotal = rs("contentTotal")
			additionalTotal = rs("additionalTotal")
		else
			calculateTotal = rs("calculateTotal")
			browseTotal = rs("browseTotal")
			loginTotal = rs("loginTotal")
			shareTotal = rs("shareTotal")
			contentTotal = rs("contentTotal")
			additionalTotal = rs("additionalTotal")
		end if
	end if 
	rs.close
	set rs = nothing
	%>
  <tr>
		<td><%=calculateTotal%></td>
    <td><a href="KpiBrowseTotal.asp?memberid=<%=memberid%>&thisyear=<%=thisyear%>"><%=browseTotal%></a></td>
    <td><a href="KpiLoginTotal.asp?memberid=<%=memberid%>&thisyear=<%=thisyear%>"><%=loginTotal%></a></td>
    <td><a href="KpiShareTotal.asp?memberid=<%=memberid%>&thisyear=<%=thisyear%>"><%=shareTotal%></a></td>
    <td><%=contentTotal%></td>
    <td><%=additionalTotal%></td>
  </tr>
  </TABLE>
	<table cellspacing="0" id="ListTable">
	<caption>互動行為詳細得點記錄</caption>
  <tr>
    <th scope="col">日期</th>
    <th width="10%" scope="col">發問</th>
    <th width="10%" scope="col">討論</th>
    <th width="10%" scope="col">評價</th>
    <th width="10%" scope="col">意見</th>
    <th width="10%" scope="col">推薦詞彙</th>
    <th width="10%" scope="col">補充解釋</th>
    <th width="10%" scope="col">態度投票</th>
	<th width="10%" scope="col">農作物地圖分享</th>
    <!--Added By Leo   2011-07-06  新增好文分享欄位-->
    <th width="10%" scope="col">好文分享</th>
	<th width="10%" scope="col">主題館評價</th>
    <!--Added By Leo   2011-07-06  新增好文分享欄位-->
  </tr>
	<%
	if thisyear = "" then
	sql = "SELECT CONVERT(varchar, shareDate, 111) AS shareDate, shareAsk, shareDiscuss, shareCommend, " & _
				"shareOpinion, shareSuggest, shareExplain, shareVote, shareJigsaw, shareRecommend,shareSubjectCommend FROM MemberGradeShare " & _
				"WHERE (memberId = '" & memberid & "') AND CONVERT(varchar, shareDate, 111) < CONVERT(varchar, GETDATE(), 111) " & _
				"ORDER BY shareDate DESC"
	elseif thisyear <> "" then
	sql = "SELECT CONVERT(varchar, shareDate, 111) AS shareDate, shareAsk, shareDiscuss, shareCommend, " & _
				"shareOpinion, shareSuggest, shareExplain, shareVote, shareJigsaw, shareRecommend,shareSubjectCommend  FROM MemberGradeShare " & _
				"WHERE (memberId = '" & memberid & "') AND (DATEDIFF(YEAR, shareDate, GETDATE()) = 0) " & _
				"ORDER BY shareDate DESC"
	end if
	set rs = conn.execute(sql)
	while not rs.eof 
	%>
  <tr>
		<td><%=rs("shareDate")%></td>
    <td><%=rs("shareAsk")%></td>
    <td><%=rs("shareDiscuss")%></td>
    <td><%=rs("shareCommend")%></td>
    <td><%=rs("shareOpinion")%></td>
    <td><%=rs("shareSuggest")%></td>
		<td><%=rs("shareExplain")%></td>
		<td><%=rs("shareVote")%></td>
		<td><%=rs("shareJigsaw")%></td>
        <!--Added By Leo   2011-07-06  好文分享欄位加入資料-->
        <td><%=rs("shareRecommend")%></td>
		<td><%=rs("shareSubjectCommend")%></td>
        <!--Added By Leo   2011-07-06  好文分享欄位加入資料-->
  </tr>
	<%
		rs.movenext
	wend
	rs.close
	set rs = nothing
	%>
	</TABLE>
	</form>  
</body>
</html>                                 
