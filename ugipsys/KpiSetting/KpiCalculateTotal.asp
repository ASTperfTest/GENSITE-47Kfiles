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
	<!--link rel="stylesheet" href="/inc/setstyle.css"-->
	<link type="text/css" rel="stylesheet" href="/css/layout.css">
	<link type="text/css" rel="stylesheet" href="/css/setstyle.css">
	<link type="text/css" rel="stylesheet" href="/css/list.css">
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
		<th scope="col" width="10%">&nbsp;</th>
		<th scope="col" width="7%">得點</th>
		<th scope="col">詳細資料</th>
	</tr>
	<%
	Dim calculateTotal, browseTotal, loginTotal, shareTotal, contentTotal, additionalTotal
	Dim browseAvg,commendAvg,discussAvg,thisYearBrowseGrade,thisYearCommendGrade,thisYearDiscussGrade
	if thisyear = "" then
	sql = "SELECT * FROM MemberGradeSummary "
	sql = sql & " WHERE MemberGradeSummary.memberId = '" & memberid & "'"
	elseif thisyear <> "" then
	sql = "SELECT isnull(GradeBrowse_thisyear.GradeBrowse, 0) as GradeBrowse, isnull(GradeLogin_thisyear.GradeLogin, 0) as GradeLogin, "
	sql = sql & " isnull(GradeShare_thisyear.GradeShare, 0) as GradeShare, Member.account as memberId , "
	sql = sql & " ISNULL(VCB.ThisYearBrowseGrade,0) AS ThisYearBrowseGrade , "
	sql = sql & " ISNULL(VCC.ThisYearCommendGrade,0) AS ThisYearCommendGrade , "
	sql = sql & " ISNULL(VCD.ThisYearDiscussGrade,0) AS ThisYearDiscussGrade , "
	sql = sql & " ISNULL(VCB.BrowseAvg ,0) AS BrowseAvg  , "
	sql = sql & " ISNULL(VCC.CommendAvg ,0) AS CommendAvg  , "
	sql = sql & " ISNULL(VCD.DiscussAvg ,0) AS DiscussAvg  , "
	sql = sql & " ISNULL(MGCBY.contentIQPoint,0) AS contentIQPoint,"
	sql = sql & " ISNULL(MGCBY.contentIQGrade,0) AS contentIQGrade,"
	sql = sql & " ISNULL(MGCBY.contentRaisePoint,0) AS contentRaisePoint,"
	sql = sql & " ISNULL(MGCBY.contentRaiseGrade,0) AS contentRaiseGrade,"
	sql = sql & " ISNULL(MGCBY.contentLogoPoint,0) AS contentLogoPoint,"
	sql = sql & " ISNULL(MGCBY.contentPlantLogPoint,0) AS contentPlantLogPoint,"
	sql = sql & " ISNULL(MGCBY.contentFishBowlPoint,0) AS contentFishBowlPoint,"
	sql = sql & " ISNULL(MGCBY.contentFishBowlGrade,0) AS contentFishBowlGrade,"
	sql = sql & " ISNULL(MGCBY.contentDrPoint,0) AS contentDrPoint,"
	sql = sql & " ISNULL(MGCBY.contentDrGrade,0) AS contentDrGrade,"
	sql = sql & " ISNULL(MGCBY.contentKnowledgeQAPoint,0) AS contentKnowledgeQAPoint,"
	sql = sql & " ISNULL(MGCBY.contentKnowledgeQAGrade,0) AS contentKnowledgeQAGrade,"
	sql = sql & " ISNULL(MGCBY.contentDr2010Point,0) AS contentDr2010Point,"
	sql = sql & " ISNULL(MGCBY.contentDr2010Grade,0) AS contentDr2010Grade,"
	sql = sql & " ISNULL(MGCBY.contentTreasureHuntPoint,0) AS contentTreasureHuntPoint,"
	sql = sql & " ISNULL(MGCBY.contentTreasureHuntGrade,0) AS contentTreasureHuntGrade,"
	'MemberGradeContentByYear 
	sql = sql & " (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0) + "
	sql = sql & " ISNULL(MGCBY.contentIQPoint,0)+ISNULL(MGCBY.contentRaisePoint,0)+ "  
	sql = sql & " ISNULL(MGCBY.contentLogoPoint,0)+ ISNULL(MGCBY.contentPlantLogPoint,0)+ "  
	sql = sql & " ISNULL(MGCBY.contentFishBowlPoint,0)+ ISNULL(MGCBY.contentDrPoint,0)+  "  
	sql = sql & " ISNULL(MGCBY.contentKnowledgeQAPoint,0)+ISNULL(MGCBY.contentDr2010Point,0)+ "  
	sql = sql & " ISNULL(MGCBY.contentTreasureHuntPoint,0) ) as contentTotal,"
	
	sql = sql & " isnull(VIEWADDTHIS.ThisYearADDpoint, 0) as additionalTotal ,"
	sql = sql & " (isnull(GradeBrowse_thisyear.GradeBrowse, 0) * 0.15 + isnull(GradeLogin_thisyear.GradeLogin, 0) * 0.2) + isnull(GradeShare_thisyear.GradeShare, 0) * 0.3 "
	sql = sql & " + ISNULL(VIEWConTHIS.ThisYearCpoint,0) * 0.2 + isnull(VIEWADDTHIS.ThisYearADDpoint, 0) * 0.15 AS total "
	
	sql = sql & " FROM Member "
	sql = sql & " left JOIN GradeLogin_thisyear  ON Member.account = GradeLogin_thisyear.memberId "
	sql = sql & " left JOIN GradeShare_thisyear  ON Member.account = GradeShare_thisyear.memberId "
	sql = sql & " left JOIN GradeBrowse_thisyear ON Member.account = GradeBrowse_thisyear.memberId "
    sql = sql & " left JOIN MemberGradeSummary B ON Member.account = B.memberId  "
	sql = sql & " LEFT JOIN GradeContentBrowse_thisyear VCB ON Member.account = VCB.ieditor  "
	sql = sql & " LEFT JOIN GradeContentCommend_thisyear VCC ON Member.account = VCC.ieditor  "
	sql = sql & " LEFT JOIN GradeContentDiscuss_thisyear VCD ON Member.account = VCD.ieditor "
	sql = sql & " LEFT JOIN View_MemberGradeAdditional_ThisYear VIEWADDTHIS ON Member.account = VIEWADDTHIS.memberId "
	sql = sql & " LEFT JOIN View_MemberGradeContent_ThisYear VIEWConTHIS ON Member.account = VIEWConTHIS.memberId "
	sql = sql & " LEFT JOIN dbo.MemberGradeContentByYear MGCBY ON Member.account = MGCBY.memberId AND (CONVERT(INT, CONVERT(varchar(4), getdate(), 120 ) ) - ISNULL(years,0) = 0) "
	sql = sql & " where Member.account ='"& memberid &"'"
	'response.write sql
	'response.end
	end if
	set rs = conn.execute(sql)
	if not rs.eof then
		if thisyear <> "" then '今年度
			calculateTotal = CInt(rs("Total"))
			browseTotal = rs("GradeBrowse")
			loginTotal = rs("GradeLogin")
			shareTotal = rs("GradeShare")
			contentTotal = rs("contentTotal")
			additionalTotal = rs("additionalTotal")
			
			browseAvg = rs("BrowseAvg")
			commendAvg = rs("CommendAvg")
			discussAvg = rs("DiscussAvg")
			thisYearBrowseGrade = rs("ThisYearBrowseGrade")
			thisYearCommendGrade = rs("ThisYearCommendGrade")
			thisYearDiscussGrade = rs("ThisYearDiscussGrade")
			
			'新增遊戲欄位
			contentIQPoint=rs("contentIQPoint")
			contentIQGrade=rs("contentIQGrade")
			contentRaisePoint=rs("contentRaisePoint")
			contentRaiseGrade=rs("contentRaiseGrade")
			contentLogoPoint=rs("contentLogoPoint")
			contentPlantLogPoint=rs("contentPlantLogPoint")
			contentFishBowlPoint=rs("contentFishBowlPoint")
			contentFishBowlGrade=rs("contentFishBowlGrade")
			contentDrPoint=rs("contentDrPoint")
			contentDrGrade=rs("contentDrGrade")
			contentKnowledgeQAPoint=rs("contentKnowledgeQAPoint")
			contentKnowledgeQAGrade=rs("contentKnowledgeQAGrade")
			contentDr2010Point=rs("contentDr2010Point")
			contentDr2010Grade=rs("contentDr2010Grade")
			contentTreasureHuntPoint=rs("contentTreasureHuntPoint")
			contentTreasureHuntGrade=rs("contentTreasureHuntGrade")
			
			
		else	'總年度
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
	  <th>會員總得分</th>
    <td><%=calculateTotal%></td>
	  <td>&nbsp;</td>
  </tr>                  
  <tr>
		<th>瀏覽行為得點</th>
    <td><a href="KpiBrowseTotal.asp?memberid=<%=memberid%>&thisyear=<%=thisyear%>"><%=browseTotal%></a></td>
    <td>
			<table cellspacing="0" id="ListdataTable">
      <tr>
				<th scope="col" width="20%">入口網單元</th>
        <th scope="col" width="20%">主題館單元</th>
        <th scope="col" width="20%">知識庫單元</th>
        <th scope="col" width="20%">知識家單元</th>
        <th scope="col">小百科單元</th>
      </tr>
			<%
				sql = "SELECT SUM(browseInterCP) AS inter, SUM(browseTopicCP) AS topic, SUM(browseCatTreeCP) AS tank, " & _
							"SUM(browseQuestionCP) + SUM(browseDiscussLP) + SUM(browseDiscussCP) AS knowledge, " & _
							"SUM(browsePediaWordCP) + SUM(browsePediaExplainLP) + SUM(browsePediaExplainCP) AS pedia " & _
							"FROM MemberGradeBrowse WHERE (memberId = '" & memberid & "') "
				if thisyear <> "" then
					sql = sql & " AND  (DATEDIFF(YEAR, browseDate, GETDATE()) = 0)"
				else
					sql = sql & " AND CONVERT(varchar, browseDate, 111) < CONVERT(varchar, GETDATE(), 111) "
				end if
				set brs = conn.execute(sql)
				if not brs.eof then
			%>
      <tr>
        <td><%=brs("inter")%></td>
        <td><%=brs("topic")%></td>
        <td><%=brs("tank")%></td>
        <td><%=brs("knowledge")%></td>
        <td><%=brs("pedia")%></td>
      </tr>
			<%
				end if
				brs.close
				set brs = nothing
			%>
      </TABLE>
		</td>
  </tr>
  <tr>
		<th>登入行為得點</th>
    <td><a href="KpiLoginTotal.asp?memberid=<%=memberid%>&thisyear=<%=thisyear%>"><%=loginTotal%></a></td>
    <td>
			<table cellspacing="0" id="ListdataTable">
      <tr>
        <th scope="col" width="14%">會員登入</th>
        <th scope="col" width="14%">IQ大挑戰</th>
        <th scope="col" width="14%">虛擬種植</th>
        <th scope="col" width="14%">生日卡片連結登入</th>
        <th scope="col" width="14%">邀請會員給點</th>
        <th scope="col" width="14%">被邀請給點</th>
        <th scope="col">周頻率給點</th>
      </tr>
			<%
				sql = "SELECT SUM(loginInterGrade) AS loginInterGrade, SUM(loginIQGrade) AS loginIQGrade, " & _
							"SUM(loginRaiseGrade) AS loginRaiseGrade, SUM(loginBirthdayGrade) AS loginBirthdayGrade, " & _
                            "SUM(loginAdditional) AS loginAdditional, SUM(loginInvite) AS loginInvite, SUM(loginIsInvited) AS loginIsInvited " & _
							"FROM MemberGradeLogin WHERE (memberId = '" & memberid & "') "
				if thisyear <> "" then
					sql = sql & " AND  (DATEDIFF(YEAR, loginDate, GETDATE()) = 0)"
				else
					sql = sql & " AND CONVERT(varchar, loginDate, 111) < CONVERT(varchar, GETDATE(), 111) "
				end if
				set brs = conn.execute(sql)
				if not brs.eof then
			%>
      <tr>
				<td><%=brs("loginInterGrade")%></td>
        <td><%=brs("loginIQGrade")%></td>
        <td><%=brs("loginRaiseGrade")%></td>
        <td><%=brs("loginBirthdayGrade")%></td>
        <td><%=brs("loginInvite")%></td>
        <td><%=brs("loginIsInvited")%></td>
        <td><%=brs("loginAdditional")%></td>
      </tr>
			<%
				end if
				brs.close
				set brs = nothing
			%>
      </TABLE>
		</td>
  </tr>
  <tr>
    <th>互動行為得點</th>
    <td><a href="KpiShareTotal.asp?memberid=<%=memberid%>&thisyear=<%=thisyear%>"><%=shareTotal%></a></td>
    <td>
			<table cellspacing="0" id="ListdataTable">
      <tr>
        <th scope="col" width="8%">發問</th>
        <th scope="col" width="8%">討論</th>
        <th scope="col" width="8%">評價</th>
        <th scope="col" width="8%">意見</th>
        <th scope="col" width="8%">推薦詞彙</th>
        <th scope="col" width="8%">補充解釋</th>
        <th scope="col" width="8%">態度投票</th>
		<th scope="col" width="8%">農作物地圖分享</th>
        <th scope="col"  width="8%">好文分享</th>
		<th scope="col"  width="8%">主題館評價</th>
      </tr>
			<%
				sql = "SELECT SUM(shareAsk) AS shareAsk, SUM(shareDiscuss) AS shareDiscuss, SUM(shareCommend) AS shareCommend, " & _
							"SUM(shareOpinion) AS shareOpinion, SUM(shareSuggest) AS shareSuggest, SUM(shareExplain) AS shareExplain, " & _
							"SUM(shareVote) AS shareVote, SUM(shareJigsaw) AS shareJigsaw, SUM(shareRecommend) AS shareRecommend,SUM(shareSubjectCommend) AS shareSubjectCommend FROM MemberGradeShare WHERE (memberId = '" & memberid & "') "
				if thisyear <> "" then
					sql = sql & " AND  (DATEDIFF(YEAR, shareDate, GETDATE()) = 0)"
				else
					sql = sql & " AND CONVERT(varchar, shareDate, 111) < CONVERT(varchar, GETDATE(), 111) "
				end if
				set brs = conn.execute(sql)
				if not brs.eof then
			%>
      <tr>
        <td><%=brs("shareAsk")%></td>
        <td><%=brs("shareDiscuss")%></td>
        <td><%=brs("shareCommend")%></td>
        <td><%=brs("shareOpinion")%></td>
        <td><%=brs("shareSuggest")%></td>
        <td><%=brs("shareExplain")%></td>
        <td><%=brs("shareVote")%></td>
		<td><%=brs("shareJigsaw")%></td>
        <td><%=brs("shareRecommend")%></td>
		<td><%=brs("shareSubjectCommend")%></td>
      </tr>
			<%
				end if
				brs.close
				set brs = nothing
			%>
      </TABLE>
		</td>
  </tr>
  <tr>
    <th>內容價值得點</th>
    <td><%=contentTotal%></td>
    <td>
			<table cellspacing="0" id="ListdataTable">
			<tr>
				<th scope="col">&nbsp;</th>
				<th width="17%" scope="col">IQ大挑戰</th>
				<th width="17%" scope="col">虛擬種植</th>
				<th width="17%" scope="col">平均(被)評價</th>
				<th width="17%" scope="col">問題平均(被)討論</th>
				<th width="17%" scope="col">問題平均(被)瀏覽</th>
				<th width="17%" scope="col">Logo競賽</th>
				<th width="17%" scope="col">種植競賽</th>
				<th width="17%" scope="col">豆仔水族箱</th>
				<th width="17%" scope="col">農業博識王</th>
				<th width="17%" scope="col">知識問答你我他</th>
				<th width="17%" scope="col">全能農知博識王</th>
				<th width="17%" scope="col">知識尋寶總動員</th>
				<th width="17%" scope="col">農情濃情拼圖樂</th>
				<th width="17%" scope="col">愛拼才會贏</th>
			</tr>
			<%
				if thisyear = "" then
					sql = "SELECT SUM(contentIQPoint) as contentIQPoint" & _
						  " ,SUM(contentIQGrade) as contentIQGrade" & _
						  " ,SUM(contentRaisePoint) as contentRaisePoint" & _
						  " ,SUM(contentRaiseGrade) as contentRaiseGrade" & _
						  " ,SUM(contentCommendPoint) as contentCommendPoint" & _
						  " ,SUM(contentCommendGrade) as contentCommendGrade" & _
						  " ,SUM(contentDiscussPoint) as contentDiscussPoint" & _
						  " ,SUM(contentDiscussGrade) as contentDiscussGrade" & _
						  " ,SUM(contentBrowsePoint) as contentBrowsePoint" & _
						  " ,SUM(contentBrowseGrade) as contentBrowseGrade" & _
						  " ,SUM(contentLogoPoint) as contentLogoPoint" & _
						  " ,SUM(contentPlantLogPoint) as contentPlantLogPoint" & _
						  " ,SUM(contentFishBowlPoint) as contentFishBowlPoint" & _
						  " ,SUM(contentFishBowlGrade) as contentFishBowlGrade" & _
						  " ,SUM(contentDrPoint) as contentDrPoint" & _
						  " ,SUM(contentDrGrade) as contentDrGrade" & _
						  " ,SUM(contentKnowledgeQAPoint) as contentKnowledgeQAPoint" & _
						  " ,SUM(contentKnowledgeQAGrade) as contentKnowledgeQAGrade" & _
						  " ,SUM(contentDr2010Point) as contentDr2010Point" & _
						  " ,SUM(contentDr2010Grade) as contentDr2010Grade" & _
						  " ,SUM(contentTreasureHuntPoint) as contentTreasureHuntPoint" & _
						  " ,SUM(contentTreasureHuntGrade) as contentTreasureHuntGrade" & _
						  " ,SUM(contentHistoryPicturePoint) as contentHistoryPicturePoint" & _
						  " ,SUM(contentPuzzle2011Point) as contentPuzzle2011Point" & _
						" FROM MemberGradeContentByYear WHERE (memberId = '" & memberid & "')"
					set brs = conn.execute(sql)
				
					if not brs.eof then
			%>
			<tr>
				<th>等級</th>
				<td><%=brs("contentIQGrade")%></td>
				<td><%=brs("contentRaiseGrade")%></td>
				<td><%=brs("contentCommendGrade")%></td>
				<td><%=brs("contentDiscussGrade")%></td>
				<td><%=brs("contentBrowseGrade")%></td>
				<td>-</td>
				<td>-</td>
				<td><%=brs("contentFishBowlGrade")%></td>
				<td><%=brs("contentDrGrade")%></td>
				<td><%=brs("contentKnowledgeQAGrade")%></td>
				<td><%=brs("contentDr2010Grade")%></td>
				<td><%=brs("contentTreasureHuntGrade")%></td>
				<td><%=brs("contentHistoryPicturePoint")%></td>
				<td><%=brs("contentPuzzle2011Point")%></td>
			</tr>
			<tr>
				<th>得點</th>
				<td><%=brs("contentIQPoint")%></td>
				<td><%=brs("contentRaisePoint")%></td>
				<td><%=brs("contentCommendPoint")%></td>
				<td><%=brs("contentDiscussPoint")%></td>
				<td><%=brs("contentBrowsePoint")%></td>
				<td><%=brs("contentLogoPoint")%></td>
				<td><%=brs("contentPlantLogPoint")%></td>
				<td><%=brs("contentFishBowlPoint")%></td>
				<td><%=brs("contentDrPoint")%></td>
				<td><%=brs("contentKnowledgeQAPoint")%></td>
				<td><%=brs("contentDr2010Point")%></td>
				<td><%=brs("contentTreasureHuntPoint")%></td>
				<td>--</td>
				<td>--</td>
			</tr>
			<%
					end if
					brs.close
					set brs = nothing
				else
					sql = "SELECT SUM(contentIQPoint) as contentIQPoint" & _
						  " ,SUM(contentIQGrade) as contentIQGrade" & _
						  " ,SUM(contentRaisePoint) as contentRaisePoint" & _
						  " ,SUM(contentRaiseGrade) as contentRaiseGrade" & _
						  " ,SUM(contentCommendPoint) as contentCommendPoint" & _
						  " ,SUM(contentCommendGrade) as contentCommendGrade" & _
						  " ,SUM(contentDiscussPoint) as contentDiscussPoint" & _
						  " ,SUM(contentDiscussGrade) as contentDiscussGrade" & _
						  " ,SUM(contentBrowsePoint) as contentBrowsePoint" & _
						  " ,SUM(contentBrowseGrade) as contentBrowseGrade" & _
						  " ,SUM(contentLogoPoint) as contentLogoPoint" & _
						  " ,SUM(contentPlantLogPoint) as contentPlantLogPoint" & _
						  " ,SUM(contentFishBowlPoint) as contentFishBowlPoint" & _
						  " ,SUM(contentFishBowlGrade) as contentFishBowlGrade" & _
						  " ,SUM(contentDrPoint) as contentDrPoint" & _
						  " ,SUM(contentDrGrade) as contentDrGrade" & _
						  " ,SUM(contentKnowledgeQAPoint) as contentKnowledgeQAPoint" & _
						  " ,SUM(contentKnowledgeQAGrade) as contentKnowledgeQAGrade" & _
						  " ,SUM(contentDr2010Point) as contentDr2010Point" & _
						  " ,SUM(contentDr2010Grade) as contentDr2010Grade" & _
						  " ,SUM(contentTreasureHuntPoint) as contentTreasureHuntPoint" & _
						  " ,SUM(contentTreasureHuntGrade) as contentTreasureHuntGrade" & _
						  " ,SUM(contentHistoryPicturePoint) as contentHistoryPicturePoint" & _
						  " ,SUM(contentPuzzle2011Point) as contentPuzzle2011Point" & _
						" FROM MemberGradeContentByYear WHERE (memberId = '" & memberid & "') " & _ 
						" AND YEARS = YEAR(GETDATE())"
					set brs = conn.execute(sql)
					if not brs.eof then
			%>
			<tr>
				<th>等級</th>
				<td><%=contentIQGrade%></td>
				<td><%=contentRaiseGrade%></td>
				<td><%=commendAvg%></td>
				<td><%=discussAvg%></td>
				<td><%=browseAvg%></td>
				<td>-</td>
				<td>-</td>
				<td><%=contentFishBowlGrade%></td>
				<td><%=contentDrGrade%></td>
				<td><%=contentKnowledgeQAGrade%></td>
				<td><%=contentDr2010Grade%></td>
				<td><%=contentTreasureHuntGrade%></td>
				<td><%=brs("contentHistoryPicturePoint")%></td>
				<td><%=brs("contentPuzzle2011Point")%></td>
			</tr>
			<tr>
				<th>得點</th>
				<td><%=contentIQPoint%></td>
				<td><%=contentRaisePoint%></td>
				<td><%=thisYearCommendGrade%></td>
				<td><%=thisYearDiscussGrade%></td>
				<td><%=thisYearBrowseGrade%></td>
				<td><%=contentLogoPoint%></td>
				<td><%=contentPlantLogPoint%></td>
				<td><%=contentFishBowlPoint%></td>
				<td><%=contentDrPoint%></td>
				<td><%=contentKnowledgeQAPoint%></td>
				<td><%=contentDr2010Point%></td>
				<td><%=contentTreasureHuntPoint%></td>
				<td>--</td>
				<td>--</td>
			</tr>
			<%
				end if
					brs.close	
				end if
				
			%>
			</TABLE>
		</td>
  </tr>
  <tr>
		<th>活動參與得點</th>
    <td><%=additionalTotal%></td>
    <td>
			<table cellspacing="0" id="ListdataTable">
      <tr>
        <th scope="col">&nbsp;</th>
        <th width="17%" scope="col">問卷調查#1</th>
        <th width="17%" scope="col">知識家活動</th>
        <th width="17%" scope="col">小百科活動</th>
        <th width="17%" scope="col">虛擬種植活動</th>
        <th width="17%" scope="col">問卷調查#2</th>
		<th width="17%" scope="col">Logo競賽</th>
        <th width="17%" scope="col">種植競賽</th>
        <th width="17%" scope="col">種植投票</th>
        <th width="17%" scope="col">豆仔水族箱</th>
        <th width="17%" scope="col">2009問卷調查#1</th>
		<th width="17%" scope="col">農業博識王</th>
		<th width="17%" scope="col">知識問答你我他</th>
		<th width="17%" scope="col">全能農知博識王</th>
		<th width="17%" scope="col">知識尋寶總動員</th>
		<th width="17%" scope="col">2010問卷調查#1</th>
		<th width="17%" scope="col">農情濃情拼圖樂</th>
		<th width="17%" scope="col">愛拼才會贏</th>
      </tr>
			<%
			if thisyear = "" then
				sql = "SELECT SUM(additionalUsage)AS additionalUsage, " & _
							"SUM(additionalKnowledge)AS additionalKnowledge, " & _
							"SUM(additionalPedia)AS additionalPedia, " & _
							"SUM(additionalRaise)AS additionalRaise, " & _
							"SUM(additionalSatisfy)AS additionalSatisfy, " & _
							"SUM(additionalLogo)AS additionalLogo, " & _
							"SUM(additionalPlantLog)AS additionalPlantLog, " & _
							"SUM(additionalPlantLogVote)AS additionalPlantLogVote, " & _
							"SUM(additionalFishBowl)AS additionalFishBowl, " & _
							"SUM(additionalQuestion)AS additionalQuestion, " & _
							"SUM(additionalDr)AS additionalDr, " & _
							"SUM(additionalKnowledgeQA)AS additionalKnowledgeQA, " & _
							"SUM(additionalDr2010)AS additionalDr2010, " & _
							"SUM(additionalTreasureHunt)AS additionalTreasureHunt, " & _
							"SUM(additionalKnowledgeAdd)AS additionalKnowledgeAdd, " & _
							"SUM(additionalPediaAdd)AS additionalPediaAdd, " & _
							"SUM(additionalRaiseAdd)AS additionalRaiseAdd, " & _
							"SUM(additionalQuestion2010)AS additionalQuestion2010, " & _
							"SUM(additionalHistoryPicture)AS additionalHistoryPicture, " & _
							"SUM(additionalPuzzle2011)AS additionalPuzzle2011 " & _
					  "FROM MemberGradeAdditional WHERE (memberId = '" & memberid & "')"
				set brs = conn.execute(sql)
				if not brs.eof then
					
			%>
		<tr>
				<th>參與</th>
        <td><%=brs("additionalUsage")%></td>
        <td><%=brs("additionalKnowledge")%></td>
        <td><%=brs("additionalPedia")%></td>
        <td><%=brs("additionalRaise")%></td>
        <td><%=brs("additionalSatisfy")%></td>
		<td><%=brs("additionalLogo")%></td>
        <td><%=brs("additionalPlantLog")%></td>
        <td><%=brs("additionalPlantLogVote")%></td>
        <td><%=brs("additionalFishBowl")%></td>
        <td><%=brs("additionalQuestion")%></td>
		<td><%=brs("additionalDr")%></td>
		<td><%=brs("additionalKnowledgeQA")%></td>
        <td><%=brs("additionalDr2010")%></td>
		<td><%=brs("additionalTreasureHunt")%></td>
		<td><%=brs("additionalQuestion2010")%></td>
		<td><%=brs("additionalHistoryPicture")%></td>
		<td><%=brs("additionalPuzzle2011")%></td>
      </tr>
      <tr>
				<th>加權得點</th>
        <td>--</td>
        <td><%=brs("additionalKnowledgeAdd")%></td>
        <td><%=brs("additionalPediaAdd")%></td>
        <td><%=brs("additionalRaiseAdd")%></td>
        <td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
        <td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
      </tr>
			<%
					end if
					brs.close
					set brs = nothing
				else
					sql = "SELECT * " & _
							"FROM MemberGradeAdditional WHERE memberId = '" & memberid & "' " & _
							" AND YEARS = YEAR(GETDATE())"
					set brs = conn.execute(sql)
					if not brs.eof then
			%>
		<tr>
				<th>參與</th>
        <td><%=brs("additionalUsage")%></td>
        <td><%=brs("additionalKnowledge")%></td>
        <td><%=brs("additionalPedia")%></td>
        <td><%=brs("additionalRaise")%></td>
        <td><%=brs("additionalSatisfy")%></td>
		<td><%=brs("additionalLogo")%></td>
        <td><%=brs("additionalPlantLog")%></td>
        <td><%=brs("additionalPlantLogVote")%></td>
        <td><%=brs("additionalFishBowl")%></td>
        <td><%=brs("additionalQuestion")%></td>
		<td><%=brs("additionalDr")%></td>
		<td><%=brs("additionalKnowledgeQA")%></td>
        <td><%=brs("additionalDr2010")%></td>
		<td><%=brs("additionalTreasureHunt")%></td>
		<td><%=brs("additionalQuestion2010")%></td>
		<td><%=brs("additionalHistoryPicture")%></td>
		<td><%=brs("additionalPuzzle2011")%></td>
      </tr>
      <tr>
				<th>加權得點</th>
        <td>--</td>
        <td><%=brs("additionalKnowledgeAdd")%></td>
        <td><%=brs("additionalPediaAdd")%></td>
        <td><%=brs("additionalRaiseAdd")%></td>
        <td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
        <td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
      </tr>
			<%
					else
					%>
		<tr>
				<th>參與</th>
        <td>0</td>
        <td>0</td>
        <td>0</td>
        <td>0</td>
        <td>0</td>
		<td>0</td>
        <td>0</td>
        <td>0</td>
        <td>0</td>
        <td>0</td>
		<td>0</td>
		<td>0</td>
        <td>0</td>
		<td>0</td>
		<td>0</td>
		<td>0</td>
		<td>0</td>
      </tr>
      <tr>
				<th>加權得點</th>
        <td>--</td>
        <td>0</td>
        <td>0</td>
        <td>0</td>
        <td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>--</td>
		<td>0</td>
        <td>0</td>
		<td>0</td>
		<td>0</td>
		<td>--</td>
		<td>--</td>
      </tr>
					<%
					end if
					brs.close
					set brs = nothing
				end if
			%>	
			</TABLE>
		</td>
  </tr>
	</TABLE>
	</form>  
</body>
</html>                                 
