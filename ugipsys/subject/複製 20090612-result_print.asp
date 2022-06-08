<%@ CodePage = 65001 %>
<% Response.Expires = 0 
HTProgCap="主題館績效統計"
response.charset = "utf-8"
HTProgCode = "webgeb2"
Response.AddHeader "content-disposition","attachment; filename=subject.xls"
Response.Charset ="utf-8"
Response.ContentType = "Content-Language;content=utf-8" 
Response.ContentType = "application/vnd.ms-excel"
   %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
subjects = split(session("subjects") , ",")
%>
<html>
<head>
<title><%= HTProgCap %></title>
</head>
<body leftmargin="0" rightmargin="0" topmargin="10" marginwidth="0" marginheight="0">
 <form name="form1" action="subjectList.asp" method="post">
<center><TABLE name="subjectlist" border="1" class="eTable" cellspacing="1" cellpadding="2" width="85%">
<tr>
	<td colspan=3>主題館績效報表</td>
	<td colspan=7>日期範圍：<%=session("DateS") %> ～ <%=session("DateE") %></td>
</tr>
<tr>
	<td>項目</td>
	<td colspan=3 align="center">基本資訊</td>
	<td colspan=4 align="center">建置面</td>
	<td colspan=2 align="center">互動面</td>
</tr>
<tr>
	<td>主題名稱</td>
	<td>建置初始日期</td>
	<td>最近更新日期</td>
	<td>管理者</td>
	<td>頁面瀏覽/次</td>
	<td width = "40px">文章/篇</td>
	<td width = "40px">圖片/張</td>
	<td width = "40px">附件/篇</td>
	<td>推薦人數</td>
	<td>總得星數</td>
</tr>
<%
for i = 0 to Ubound(subjects)
	'response.write i&" : "&subjects(i) &"<br>"
	subjectid = subjects(i)
	subjectid2 = subjects(i)
	if trim(subjects(i)) = "31" then '蝴蝶主題館
		subjectid2 = 147
	end if
	if trim(subjects(i)) = "32" then '鳳梨主題館
		subjectid2 = 148
	end if
	'取主題館名稱 建置日期 管理者
	sql1 = "SELECT CatTreeRoot.CtRootName, NodeInfo.create_time,NodeInfo.owner "
	sql1 = sql1 & " FROM NodeInfo LEFT OUTER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
	sql1 = sql1 & " WHERE CatTreeRoot.CtRootID = '"& subjectid2 &"'"
	set rs1 = conn.execute(sql1)
	'取管理者資料
	sql2 = "SELECT InfoUser.UserName, Dept.deptName "
	sql2 = sql2 & " FROM InfoUser INNER JOIN Dept ON InfoUser.deptID = Dept.deptID"
	sql2 = sql2 & " WHERE InfoUser.UserID = '"& rs1("owner")&"'"
	set rs2 = conn.execute(sql2)
	'取文章資料
	'取篇數 圖片數 最近更新日期
	sql3 = "SELECT COUNT(*) as count1,COUNT(CuDTGeneric.xImgFile) as count2 , "
	sql3 = sql3 & " max(CuDTGeneric.dEditDate) as newdate  "
	' sql3 = sql3 & " SUM(CuDTGeneric.ClickCount) as sum1, "
	' sql3 = sql3 & " SUM(DailyClick.dailyClick) as sum2, " '瀏覽數2
	' sql3 = sql3 & " SUM(SubjectForum.CommandCount) as sumP , "
	' sql3 = sql3 & " SUM(SubjectForum.GradeCount) as sumS "
	sql3 = sql3 & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	'sql3 = sql3 & " LEFT JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	'sql3 = sql3 & " LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem "
	sql3 = sql3 & " WHERE CatTreeNode.CtRootID = '"& subjectid &"'"
	if session("DateS")  <> "" then
		sql3 = sql3 & " and CuDTGeneric.xPostDate >= '"& session("DateS")  &"'"
	end if
	if session("DateE")  <> "" then
		sql3 = sql3 & " and CuDTGeneric.xPostDate <= '"& session("DateE")  &"'"
	end if
	set rs3 = conn.execute(sql3)
	'取瀏覽數
	sql3_2 = "SELECT SUM(DailyClick.dailyClick) as sum2" '瀏覽數2
	sql3_2 = sql3_2 & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	'sql3_2 = sql3_2 & " LEFT JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	sql3_2 = sql3_2 & " LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem "
	sql3_2 = sql3_2 & " WHERE CatTreeNode.CtRootID = '"& subjectid &"'"
	if session("DateS") <> "" then
		sql3_2 = sql3_2 & " and DailyClick.editDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql3_2 = sql3_2 & " and DailyClick.editDate <= '"& session("DateE") &"'"
	end if
	set rs3_2 = conn.execute(sql3_2)
	'取推薦人數 總得星數
	sql3_3 = "SELECT "
	sql3_3 = sql3_3 & " SUM(SubjectForum.CommandCount) as sumP , " '推薦人數
	sql3_3 = sql3_3 & " SUM(SubjectForum.GradeCount) as sumS " '總得星數
	sql3_3 = sql3_3 & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	sql3_3 = sql3_3 & " LEFT JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	'sql3 = sql3 & " LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem "
	sql3_3 = sql3_3 & " WHERE CatTreeNode.CtRootID = '"& subjectid &"'"
	if session("DateS") <> "" then
		sql3_3 = sql3_3 & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql3_3 = sql3_3 & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	set rs3_3 = conn.execute(sql3_3)
	'取附件數
	sql4 = "SELECT COUNT(*) as count1 "
	sql4 = sql4 & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	sql4 = sql4 & " INNER JOIN CuDTAttach ON CuDTGeneric.iCUItem = CuDTAttach.xiCuItem "
	sql4 = sql4 & " WHERE CatTreeNode.CtRootID = '"& subjectid &"'"
	if session("DateS")  <> "" then
		sql4 = sql4 & " and CuDTGeneric.xPostDate >= '"& session("DateS")  &"'"
	end if
	if session("DateE")  <> "" then
		sql4 = sql4 & " and CuDTGeneric.xPostDate <= '"& session("DateE")  &"'"
	end if
	set rs4 = conn.execute(sql4)
	'response.write sql3
%>
<tr>
	<td><%=subjects(i)%>.<%=rs1("CtRootName")%></td>
	<td><%=datevalue(rs1("create_time"))%></td>
	<td><% '最近更新日期
		if isnull(rs3("newdate")) then
			response.write "&nbsp;"
		else
			response.write datevalue(rs3("newdate"))
		end if
	%></td>
	<td><%if not rs2.eof then
					response.write rs2("UserName") & " / " & rs2("deptName")
				end if
	%>&nbsp;</td>
	<td><%
		if isnull(rs3_2("sum2")) then '頁面瀏覽次數
			response.write "0"
		else
			response.write rs3_2("sum2")
		end if
	%></td>
	<td><%=rs3("count1")%></td><!--文章-->
	<td><%=rs3("count2")%></td><!--圖片-->
	<td><%=rs4("count1")%></td><!--附件-->
	<td><%
		if isnull(rs3_3("sumP")) then
			response.write "0"
		else
			response.write rs3_3("sumP")
		end if
	%></td><!--推薦人數-->
	<td><%
		if isnull(rs3_3("sumS")) then
			response.write "0"
		else
			response.write rs3_3("sumS")
		end if
	%></td><!--總得星數-->
</tr>
<%
next
%>
</TABLE>
</center>
</form>