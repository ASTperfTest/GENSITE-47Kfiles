<%@ CodePage = 65001 %>
<% Response.Expires = 0 
HTProgCap="主題館績效統計"
response.charset = "utf-8"
HTProgCode = "webgeb2"
   %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
' if request("pcShowhtx_IDateS") = "" and request("pcShowhtx_IDateE") = "" then
	' Response.Write "<script language='javascript'>"
	' Response.Write "alert('請選擇日期範圍！');"
	' Response.Write "history.go(-1);"
	' Response.Write "</script>"
' end if
if request("ckbox") <> "" then
	session("subjects") = request("ckbox")
	session("subjects") = replace(session("subjects"),"147","31")
	session("subjects") = replace(session("subjects"),"148","32")
	session("Osubjects") = session("subjects") '原本選的主題館
end if
if session("subjects") = "" then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('請至少選擇一個主題館！');"
	Response.Write "history.go(-1);"
	Response.Write "</script>"
end if

if request("pcShowhtx_IDateS") <> "" then
	session("DateS") = request("pcShowhtx_IDateS")
end if
if request("pcShowhtx_IDateE") <> "" then
	session("DateE") = request("pcShowhtx_IDateE")
end if
response.write "<hr>"
'排序主題館
sql = ""
if request("orderby") = "click" then
	sql = "SELECT CatTreeNode.CtRootID,isnull(SUM(DailyClick.dailyClick),0) as click" '瀏覽數
	sql = sql & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	'sql = sql & " LEFT JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	sql = sql & " LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem "
	sql = sql & " WHERE CatTreeNode.CtRootID in ("& session("subjects") &") "
	if session("DateS") <> "" then
		sql = sql & " and DailyClick.editDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql = sql & " and DailyClick.editDate <= '"& session("DateE") &"'"
	end if
	sql = sql & " GROUP BY CatTreeNode.CtRootID "
	sql = sql & " ORDER BY click DESC , CatTreeNode.CtRootID"
elseif request("orderby") = "article" then
	sql = "SELECT CatTreeNode.CtRootID ,isnull(COUNT(*),0) AS article "
	sql = sql & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	'sql = sql & " LEFT JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	'sql = sql & " LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem "
	sql = sql & " WHERE CatTreeNode.CtRootID in ("& session("subjects") &") "
	if session("DateS") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	sql = sql & " GROUP BY CatTreeNode.CtRootID "
	sql = sql & " ORDER BY article DESC , CatTreeNode.CtRootID"
elseif request("orderby") = "pic" then
	sql = "SELECT CatTreeNode.CtRootID ,isnull(COUNT(CuDTGeneric.xImgFile),0) as pic "
	sql = sql & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	'sql = sql & " LEFT JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	'sql = sql & " LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem "
	sql = sql & " WHERE CatTreeNode.CtRootID in ("& session("subjects") &")"
	if session("DateS") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	sql = sql & " GROUP BY CatTreeNode.CtRootID "
	sql = sql & " ORDER BY pic DESC , CatTreeNode.CtRootID"
elseif request("orderby") = "person" then
	sql = "SELECT CatTreeNode.CtRootID ,isnull(SUM(SubjectForum.CommandCount),0) as sumP "
	sql = sql & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	sql = sql & " LEFT JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	'sql = sql & " LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem "
	sql = sql & " WHERE CatTreeNode.CtRootID in ("& session("subjects") &")"
	if session("DateS") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	sql = sql & " GROUP BY CatTreeNode.CtRootID "
	sql = sql & " ORDER BY sumP DESC , CatTreeNode.CtRootID"
elseif request("orderby") = "star" then
	sql = "SELECT CatTreeNode.CtRootID ,isnull(SUM(SubjectForum.GradeCount),0) as sumS "
	sql = sql & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	sql = sql & " LEFT JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	'sql = sql & " LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem "
	sql = sql & " WHERE CatTreeNode.CtRootID in ("& session("subjects") &")"
	if session("DateS") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	sql = sql & " GROUP BY CatTreeNode.CtRootID "
	sql = sql & " ORDER BY sumS DESC , CatTreeNode.CtRootID"
elseif request("orderby") = "attach" then
	sql = "SELECT CatTreeNode.CtRootID,isnull(COUNT(*),0) as count1 "
	'sql = sql & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	'sql = sql & " INNER JOIN CuDTAttach ON CuDTGeneric.iCUItem = CuDTAttach.xiCuItem "
	sql = sql & " FROM CuDTAttach INNER JOIN CuDTGeneric ON CuDTAttach.xiCuItem = CuDTGeneric.iCUItem "
	sql = sql & " RIGHT OUTER JOIN CatTreeNode ON CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID "
					  
	sql = sql & " WHERE CatTreeNode.CtRootID in ("& session("subjects") &")"
	if session("DateS") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	sql = sql & " GROUP BY CatTreeNode.CtRootID "
	sql = sql & " ORDER BY count1 DESC , CatTreeNode.CtRootID"
end if
if sql <> "" then
	'response.write sql
	set rs = conn.execute(sql)
	session("subjects") = ""
	while not rs.eof
		session("subjects") = session("subjects") &","&rs("CtRootID")
		rs.movenext
	wend
	session("subjects") = RIGHT(session("subjects"),LEN(session("subjects"))-1)
end if
subjects = split(session("subjects") , ",")
'排序主題館end
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
<center><TABLE class="ListTable" name="subjectlist" border="1" cellspacing="0" width="85%">
<tr>
	<td colspan=2>主題館績效報表</td>
	<td colspan=5>日期範圍：<%=session("DateS")%> ～ <%=session("DateE")%></td>
	<td colspan=3>
	<select name="orderby" id="orderby" >	
		<option value="" <%if request("orderby")="" then%>selected<%END IF%>>請選擇</option>
		<option value="click" <%if request("orderby")="click" then%>selected<%END IF%>>頁面瀏覽/次</option>
		<option value="article" <%if request("orderby")="article" then%>selected<%END IF%>>文章/篇</option>
		<option value="pic" <%if request("orderby")="pic" then%>selected<%END IF%>>圖片/張</option>
		<option value="attach" <%if request("orderby")="attach" then%>selected<%END IF%>>附件/篇</option>
		<option value="person" <%if request("orderby")="person" then%>selected<%END IF%>>推薦人數</option>
		<option value="star" <%if request("orderby")="star" then%>selected<%END IF%>>總得星數</option>
	</select>
	</td>
</tr>
<tr>
	<td>項目</td>
	<td colspan=3 align="center">基本資訊</td>
	<td colspan=4 align="center">建置面</td>
	<td colspan=2 align="center">互動面</td>
</tr>
<tr>
	<th  width = "15%">主題館名稱</th>
	<th scope="col">建置初始日期</th>
	<th>最近更新日期</th>
	<th>管理者</th>
	<th>頁面瀏覽/次</th>
	<th width = "40px">文章/篇</th>
	<th width = "40px">圖片/張</th>
	<th width = "40px">附件/篇</th>
	<th>推薦人數</th>
	<th>總得星數</th>
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
	sql3 = "SELECT COUNT(*) as count1,COUNT(CuDTGeneric.xImgFile) as count2 ," '圖片數
	sql3 = sql3 & " max(CuDTGeneric.dEditDate) as newdate " '更新日期
	'sql3 = sql3 & " SUM(CuDTGeneric.ClickCount) as sum1, " '瀏覽數
	'sql3 = sql3 & " SUM(DailyClick.dailyClick) as sum2, " '瀏覽數2
	'sql3 = sql3 & " SUM(SubjectForum.CommandCount) as sumP , " '推薦人數
	'sql3 = sql3 & " SUM(SubjectForum.GradeCount) as sumS " '總得星數
	sql3 = sql3 & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	'sql3 = sql3 & " LEFT JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	'sql3 = sql3 & " LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem "
	sql3 = sql3 & " WHERE CatTreeNode.CtRootID = '"& subjectid &"'"
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
	if session("DateS") <> "" then
		sql4 = sql4 & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql4 = sql4 & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	set rs4 = conn.execute(sql4)
	'response.write sql3
%>
<tr>
	<td><a href = "result_subject.asp?id=<%=trim(subjectid2)%>"><%=rs1("CtRootName")%></a></td>
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
	<td><a href = "result_article.asp?atype=1&id=<%=trim(subjects(i))%>"><%=rs3("count1")%></a></td><!--文章-->
	<td><a href = "result_article.asp?atype=2&id=<%=trim(subjects(i))%>"><%=rs3("count2")%></a></td><!--圖片-->
	<td><a href = "result_article.asp?atype=3&id=<%=trim(subjects(i))%>"><%=rs4("count1")%></a></td><!--附件-->
	<td><a href = "result_article.asp?atype=4&id=<%=trim(subjects(i))%>"><%
		if isnull(rs3_3("sumP")) then
			response.write "0"
		else
			response.write rs3_3("sumP")
		end if
	%></a></td><!--推薦人數-->
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
<tr><td colspan=10 align=center>
	<input type="button" class="cbutton" name="excel" value="匯出Excel" onclick = "javascript:location.href='result_print.asp'">
	<input type="button" class="cbutton" name="research" value="重新查詢" onclick = "javascript:location.href='subjectList.asp?subjectid=all'">
	</td>
</tr>
</TABLE>
</center>
</form>
<script Language=VBScript>
	sub orderby_OnChange
		document.location.href="result.asp?orderby="&form1.orderby.value
	 end sub
</script>