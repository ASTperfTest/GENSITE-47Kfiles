<%@ CodePage = 65001 %>
<% Response.Expires = 0 
HTProgCap="主題館績效統計"
response.charset = "utf-8"
HTProgCode = "webgeb2"
   %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
id =request("id")
response.write "<hr>"
%>
<html>
<head>
<title><%= HTProgCap %></title>
<link href="../css/list.css" rel="stylesheet" type="text/css">
<link href="../css/layout.css" rel="stylesheet" type="text/css">
<link href="../css/form.css" rel="stylesheet" type="text/css" />
</head>
<body leftmargin="0" rightmargin="0" topmargin="10" marginwidth="0" marginheight="0">
 <form name="form1" action="subjectList.asp" method="post">
<center><TABLE name="subjectlist" border="1" class="eTable" cellspacing="0" cellpadding="2" width="80%">
<tr>
	<td colspan=2>主題館績效報表</td>
	<td colspan=6>日期範圍：<%=session("DateS")%> ～ <%=session("DateE")%></td>
</tr>
<tr>
	<td>項目</td>
	<td colspan=1 align="center">基本資訊</td>
	<td colspan=4 align="center">建置面</td>
	<td colspan=2 align="center">互動面</td>
</tr>
<tr>
	<td>主題館名稱</td>
	<td>單位</td>
	<td>頁面瀏覽/次</td>
	<td>文章/篇</td>
	<td>圖片/張</td>
	<td>附件/篇</td>
	<td>推薦人數</td>
	<td>總得星數</td>
</tr>
<%
	if trim(id) = "147" then '蝴蝶主題館
		id = 31
	end if
	if trim(id) = "148" then '鳳梨主題館
		id = 32
	end if
	'取主題館名稱
	if id =31 or id = 32 then
		sql1 = "SELECT CatTreeRoot.CtRootName "
		sql1 = sql1 & " FROM  CatTreeRoot "
		sql1 = sql1 & " WHERE CatTreeRoot.CtRootID = '"& id &"'"
	else
		sql1 = "SELECT CatTreeRoot.CtRootName "
		sql1 = sql1 & " FROM NodeInfo LEFT OUTER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID "
		sql1 = sql1 & " WHERE CatTreeRoot.CtRootID = '"& id &"'"
	end if
	set rs1 = conn.execute(sql1)
	'取上稿單位
	sql2 = "SELECT DISTINCT CuDTGeneric.iEditor, InfoUser.UserName, InfoUser.deptID, Dept.deptName "
	' sql2 = sql2 & " FROM Dept INNER JOIN InfoUser ON Dept.deptID = InfoUser.deptID "
	' sql2 = sql2 & " RIGHT OUTER JOIN "
	' sql2 = sql2 & " CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	' sql2 = sql2 & " ON InfoUser.UserID = CuDTGeneric.iEditor "
	sql2 = sql2 & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	sql2 = sql2 & " INNER JOIN InfoUser ON CuDTGeneric.iEditor = InfoUser.UserID "
	sql2 = sql2 & " INNER JOIN Dept ON InfoUser.deptID = Dept.deptID "
	sql2 = sql2 & " WHERE CatTreeNode.CtRootID = '"& id &"'"
	set rs2 = conn.execute(sql2)
	'初始化
	pageclick =0
	article =0
	pic =0
	attach =0
	C_P = 0 '推薦人數
	C_S = 0 '總得星數
while not rs2.eof
	'取文章資料
	'取篇數 圖片數 最近更新日期
	sql3 = "SELECT COUNT(*) as count1,COUNT(CuDTGeneric.xImgFile) as count2 , "
	sql3 = sql3 & " max(CuDTGeneric.dEditDate) as newdate  "
	'sql3 = sql3 & " SUM(CuDTGeneric.ClickCount) as sum1, "
	'sql3 = sql3 & " SUM(DailyClick.dailyClick) as sum2, " '瀏覽數2
	'sql3 = sql3 & " SUM(SubjectForum.CommandCount) as sumP , "
	'sql3 = sql3 & " SUM(SubjectForum.GradeCount) as sumS "
	sql3 = sql3 & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	'sql3 = sql3 & " LEFT JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	sql3 = sql3 & " LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem "
	sql3 = sql3 & " WHERE CatTreeNode.CtRootID = '"& id &"' and CuDTGeneric.iEditor = '"& rs2("iEditor")&"'"
	if session("DateS") <> "" then
		sql3 = sql3 & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql3 = sql3 & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	set rs3 = conn.execute(sql3)
	'取瀏覽數
	sql3_2 = "SELECT SUM(DailyClick.dailyClick) as sum2" '瀏覽數2
	sql3_2 = sql3_2 & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	'sql3_2 = sql3_2 & " LEFT JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	sql3_2 = sql3_2 & " LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem "
	sql3_2 = sql3_2 & " WHERE CatTreeNode.CtRootID = '"& id &"' and CuDTGeneric.iEditor = '"& rs2("iEditor")&"'"
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
	sql3_3 = sql3_3 & " WHERE CatTreeNode.CtRootID = '"& id &"' and CuDTGeneric.iEditor = '"& rs2("iEditor")&"'"
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
	sql4 = sql4 & " WHERE CatTreeNode.CtRootID = '"& id &"' and CuDTGeneric.iEditor = '"& rs2("iEditor")&"'"
	if session("DateS") <> "" then
		sql4 = sql4 & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql4 = sql4 & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	set rs4 = conn.execute(sql4)
	'response.write sql2
%>
<tr>
	<td><%=rs1("CtRootName")%></td>
	<td><%=rs2("deptName")%></td>
	<td><%'頁面瀏覽次數
		if isnull(rs3_2("sum2")) then
			response.write "0"
			pageclick =pageclick + 0
		else
			response.write rs3_2("sum2")
			pageclick =pageclick + rs3_2("sum2")
		end if
	%></td>
	<td><%=rs3("count1")%></td><!--文章-->
	<td><%=rs3("count2")%></td><!--圖片-->
	<td><%=rs4("count1")%></td><!--附件-->
	<td><%
		if isnull(rs3_3("sumP")) then
			response.write "0"
			C_P = C_P + 0
		else
			response.write rs3_3("sumP")
			C_P = C_P + rs3_3("sumP")
		end if
	%></td><!--推薦人數-->
	<td><%
		if isnull(rs3_3("sumS")) then
			response.write "0"
			C_S = C_S + 0
		else
			response.write rs3_3("sumS")
			C_S = C_S + rs3_3("sumS")
		end if
	%></td><!--總得星數-->
</tr>
<%
	article =article + rs3("count1")
	pic =pic + rs3("count2")
	attach =attach + rs4("count1")
	rs2.movenext
wend
%>
<tr>
	<td>總計</td>
	<td>&nbsp;</td>
	<td><%=pageclick%></td>
	<td><%=article%></td>
	<td><%=pic%></td>
	<td><%=attach%></td>
	<td><%=C_P%></td>
	<td><%=C_S%></td>
</tr>
</TABLE>
</center>
</form>