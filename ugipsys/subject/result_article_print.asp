<%@ CodePage = 65001 %>
<% Response.Expires = 0 
HTProgCap="主題館績效統計"
response.charset = "utf-8"
HTProgCode = "webgeb2"
Response.AddHeader "content-disposition","attachment; filename=subject_article.xls"
Response.Charset ="utf-8"
Response.ContentType = "Content-Language;content=utf-8" 
Response.ContentType = "application/vnd.ms-excel"
   %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
id =request("id")
if trim(id) = "147" then '蝴蝶主題館
	id = 31
end if
if trim(id) = "148" then '鳳梨主題館
	id = 32
end if
atype =request("atype")
if atype = "1" then
	subjecttitle = "主題館建置面(含不公開文章篇數)清單"
	
	sql = "SELECT isnull(CtUnit.CtUnitName, '') as CtUnitName"
	sql = sql & "   ,CuDTGeneric.sTitle"
	sql = sql & "   ,CuDTGeneric.xPostDate"
	sql = sql & "   ,CuDTGeneric.iCUItem"
	sql = sql & "   ,CuDTGeneric.fCTUPublic"
	sql = sql & "   ,CuDTGeneric.clickcount"
	sql = sql & "   ,isnull( Dept.deptName, '') as deptName"
	sql = sql & "   ,CatTreeNode.CtNodeID "
	sql = sql & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	sql = sql & " left JOIN InfoUser ON CuDTGeneric.iEditor = InfoUser.UserID "
	sql = sql & " left JOIN Dept ON InfoUser.deptID = Dept.deptID "
	sql = sql & " left JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID "
	sql = sql & " WHERE CatTreeNode.CtRootID = '"& id &"'" 
	if session("DateS") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	if request("orderby") = "unit" then
		sql = sql & " ORDER BY CtUnit.CtUnitName"
	elseif request("orderby") = "postdate" then
		sql = sql & " ORDER BY CuDTGeneric.xPostDate DESC "
	elseif request("orderby") = "dept" then
		sql = sql & " ORDER BY Dept.deptName"
	elseif request("orderby") = "clickcount" then
		sql = sql & " ORDER BY CuDTGeneric.clickcount"
	end if
	
	set rs = conn.execute(sql)
	
elseif atype = "5" then
	subjecttitle = "主題館建置面(文章篇數)清單"
	
	sql = "SELECT isnull(CtUnit.CtUnitName, '') as CtUnitName"
	sql = sql & "   ,CuDTGeneric.sTitle"
	sql = sql & "   ,CuDTGeneric.xPostDate"
	sql = sql & "   ,CuDTGeneric.iCUItem"
	sql = sql & "   ,CuDTGeneric.fCTUPublic"
	sql = sql & "   ,CuDTGeneric.clickcount"
	sql = sql & "   ,isnull( Dept.deptName, '') as deptName"
	sql = sql & "   ,CatTreeNode.CtNodeID "
	sql = sql & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	sql = sql & " left JOIN InfoUser ON CuDTGeneric.iEditor = InfoUser.UserID "
	sql = sql & " left JOIN Dept ON InfoUser.deptID = Dept.deptID "
	sql = sql & " left JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID "
	sql = sql & " WHERE CatTreeNode.CtRootID = '"& id &"'and CuDTGeneric.fCTUPublic = 'Y'" 
	if session("DateS") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	if request("orderby") = "unit" then
		sql = sql & " ORDER BY CtUnit.CtUnitName"
	elseif request("orderby") = "postdate" then
		sql = sql & " ORDER BY CuDTGeneric.xPostDate DESC "
	elseif request("orderby") = "dept" then
		sql = sql & " ORDER BY Dept.deptName"
	elseif request("orderby") = "clickcount" then
		sql = sql & " ORDER BY CuDTGeneric.clickcount"
	end if
	
	set rs = conn.execute(sql)

elseif atype = "2" then
	
	subjecttitle = "主題館建置面(圖片張數)清單"
	
	sql = "SELECT isnull(CtUnit.CtUnitName, '') as CtUnitName"
	sql = sql & "   ,CuDTGeneric.sTitle "
	sql = sql & "   ,CuDTGeneric.xPostDate "
	sql = sql & "   ,CuDTGeneric.iCUItem "
	sql = sql & "   ,CuDTGeneric.fCTUPublic"
	sql = sql & "   ,CuDTGeneric.clickcount"
	sql = sql & "   ,isnull( Dept.deptName, '') as deptName"
	sql = sql & "   ,CatTreeNode.CtNodeID "
	sql = sql & " FROM CatTreeNode INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	sql = sql & " left JOIN InfoUser ON CuDTGeneric.iEditor = InfoUser.UserID "
	sql = sql & " left JOIN Dept ON InfoUser.deptID = Dept.deptID "
	sql = sql & " left JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID "
	sql = sql & " WHERE CuDTGeneric.xImgFile is not NULL and CatTreeNode.CtRootID = '"& id &"'" 'and CuDTGeneric.fCTUPublic = 'Y'"   文章/篇(含不公開) 故將條件註解掉
	if session("DateS") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	if request("orderby") = "unit" then
		sql = sql & " ORDER BY CtUnit.CtUnitName"
	elseif request("orderby") = "postdate" then
		sql = sql & " ORDER BY CuDTGeneric.xPostDate DESC "
	elseif request("orderby") = "dept" then
		sql = sql & " ORDER BY Dept.deptName"
	elseif request("orderby") = "clickcount" then
		sql = sql & " ORDER BY CuDTGeneric.clickcount"
	end if
	set rs = conn.execute(sql)
	
elseif atype = "3" then
	
	subjecttitle = "主題館建置面(附件/篇)清單"
	
	sql = "SELECT DISTINCT CuDTGeneric.iCUItem"
	sql = sql & "   ,isnull(CtUnit.CtUnitName, '') as CtUnitName"
	sql = sql & "   ,CuDTGeneric.sTitle"
	sql = sql & "   ,CuDTGeneric.xPostDate"
	sql = sql & "   ,CuDTGeneric.fCTUPublic"
	sql = sql & "   ,CuDTGeneric.clickcount"
	sql = sql & "   ,isnull( Dept.deptName, '') as deptName"
	sql = sql & "   ,CatTreeNode.CtNodeID "
	sql = sql & "   ,(SELECT COUNT(*) FROM CuDTAttach WHERE xiCuItem =CuDTGeneric.iCUItem) as count "
	sql = sql & " FROM CuDTGeneric INNER JOIN CatTreeNode ON CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID "
	sql = sql & " INNER JOIN CuDTAttach ON CuDTGeneric.iCUItem = CuDTAttach.xiCuItem "
	sql = sql & " left JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID "
	sql = sql & " left JOIN InfoUser ON CuDTGeneric.iEditor = InfoUser.UserID "
	sql = sql & " left JOIN Dept ON InfoUser.deptID = Dept.deptID "
	sql = sql & " WHERE CatTreeNode.CtRootID = '"& id &"'" 'and CuDTGeneric.fCTUPublic = 'Y'"  文章/篇(含不公開) 故將條件註解掉
	if session("DateS") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	
	if request("orderby") = "unit" then
		sql = sql & " ORDER BY CtUnit.CtUnitName"
	elseif request("orderby") = "postdate" then
		sql = sql & " ORDER BY CuDTGeneric.xPostDate DESC "
	elseif request("orderby") = "dept" then
		sql = sql & " ORDER BY Dept.deptName"
	elseif request("orderby") = "attach" then
		sql = sql & " ORDER BY count"
	elseif request("orderby") = "clickcount" then
		sql = sql & " ORDER BY CuDTGeneric.clickcount"
	end if
	set rs = conn.execute(sql)

elseif atype = "4" then
	
	subjecttitle = "主題館互動面內容清單"
	
	sql = "SELECT isnull(CtUnit.CtUnitName, '') as CtUnitName"
	sql = sql & "   ,CuDTGeneric.sTitle"
	sql = sql & "   ,CuDTGeneric.xPostDate"
	sql = sql & "   ,CuDTGeneric.iCUItem"
	sql = sql & "   ,CuDTGeneric.fCTUPublic"
	sql = sql & "   ,CuDTGeneric.clickcount"
	sql = sql & "   ,isnull( Dept.deptName, '') as deptName"
	sql = sql & "   ,CatTreeNode.CtNodeID"
	sql = sql & "   ,SubjectForum.CommandCount"
	sql = sql & "   ,SubjectForum.GradeCount"
	sql = sql & "   ,SubjectForum.GradePersonCount "
	sql = sql & " FROM CatTreeNode "
	sql = sql & " INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
	sql = sql & " INNER JOIN SubjectForum ON CuDTGeneric.iCUItem = SubjectForum.gicuitem "
	sql = sql & " left JOIN InfoUser ON CuDTGeneric.iEditor = InfoUser.UserID "
	sql = sql & " left JOIN Dept ON InfoUser.deptID = Dept.deptID "
	sql = sql & " left JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID "	
	sql = sql & " WHERE SubjectForum.GradePersonCount is not null and CatTreeNode.CtRootID = '"& id &"'" 'and CuDTGeneric.fCTUPublic = 'Y'"  文章/篇(含不公開) 故將條件註解掉
	if session("DateS") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate >= '"& session("DateS") &"'"
	end if
	if session("DateE") <> "" then
		sql = sql & " and CuDTGeneric.xPostDate <= '"& session("DateE") &"'"
	end if
	if request("orderby") = "dept" then
		sql = sql & " ORDER BY Dept.deptName "
	elseif request("orderby") = "person" then
		sql = sql & " ORDER BY SubjectForum.CommandCount DESC "
	elseif request("orderby") = "star" then
		sql = sql & " ORDER BY SubjectForum.GradeCount DESC "
	elseif request("orderby") = "avgstar" then
		sql = sql & " ORDER BY (CASE WHEN SubjectForum.GradePersonCount < 1 THEN 0 ELSE SubjectForum.GradeCount / SubjectForum.GradePersonCount END) DESC "
	elseif request("orderby") = "clickcount" then
		sql = sql & " ORDER BY CuDTGeneric.clickcount"
	end if
	set rs = conn.execute(sql)
		
end if
response.write "<hr>"
%>
<html>
<head>
<title><%= HTProgCap %></title>
</head>
<body leftmargin="0" rightmargin="0" topmargin="10" marginwidth="0" marginheight="0">
 <form name="form1" action="subjectList.asp" method="post">
<center><TABLE name="subjectlist" border="1" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<tr>
	<td colspan=2><%=subjecttitle%></td>
	<%if atype = "3" then%>
	<td colspan=4>日期範圍：<%=session("DateS")%> ～ <%=session("DateE")%></td>
	<%elseif atype = "4" then%>
	<td colspan=5>日期範圍：<%=session("DateS")%> ～ <%=session("DateE")%></td>
	<%else%>
	<td colspan=3>日期範圍：<%=session("DateS")%> ～ <%=session("DateE")%></td>
	<%end if%>
</tr>
<tr>
	<td>編號</td>
	<td>標題</td>
	<td width="10%">文件瀏覽次數</td>
	<td width="10%">公開/不公開</td>
	<%if atype = "4" then%>
	<td>推薦人數</td>
	<td>總得星數</td>
	<td>平均得星</td>
	<%else%>
	<td>單元</td>
	<%end if%>
	<%if atype = "3" then%>
	<td>附件/篇</td>
	<%end if%>
	<td>張貼日期</td>
	<td>上稿單位</td>
</tr>
<%
	i=1
	
	CommandCount = 0 '推薦人數
	GradeCount = 0 '總得星數
	GradePersonCount = 0 '星數人數
	while not rs.eof
%>
<tr>
	<td><%=i%></td><!--編號-->
	<td><a align="right" target = "new" href = "http://kmw.hyweb.com.tw/subject/ct.asp?xItem=<%=rs("iCUItem")%>&ctNode=<%=rs("CtNodeID")%>&mp=<%=id%>"><%=rs("sTitle")%></a></td><!--標題-->
	<!--td><a target = "new" href = "http://kmw.hyweb.com.tw/subject/content.asp?mp=<%=id%>&CuItem=<%=rs("iCUItem")%>"><%=rs("sTitle")%></a></td-->
	<td><%=rs("clickcount")%></td><!--文件瀏覽次數-->
	<%if rs("fCTUPublic") = "Y" then%><!--公開/不公開-->
	<td>公開</td>
	<%else%>
	<td>不公開</td>
	<%end if%>
	<%if atype = "4" then
	CommandCount = CommandCount + rs("CommandCount")
	GradeCount = GradeCount + rs("GradeCount")
	GradePersonCount = GradePersonCount + rs("GradePersonCount")
	%>
	<td><%=rs("CommandCount")%></td><!--推薦人數-->
	<td><%=rs("GradeCount")%></td><!--總得星數-->
	<td><%
	if rs("GradePersonCount") <> 0 then
		response.write LEFT(rs("GradeCount")/rs("GradePersonCount"),4)
	else
		response.write 0
	end if%></td><!--平均得星-->
	<%else%>
	<td><%=rs("CtUnitName")%></td><!--單元-->
	<%end if%>
	<%if atype = "3" then%>
	<td><%=rs("count")%></td><!--附件-->
	<%end if%>
	<td><%=datevalue(rs("xPostDate"))%></td><!--張貼日期-->
	<td><%=rs("deptName")%></td><!--上稿單位-->
</tr>
<%	rs.movenext
		i=i+1
	wend
%>
<%if atype = "4" then%>
<tr>
	<td>總計</td>
	<td>&nbsp;</td>
	<td><%=CommandCount%></td>
	<td><%=GradeCount%></td>
	<td><%
	if GradePersonCount <> 0 then
		response.write LEFT(GradeCount/GradePersonCount,4)
	else
		response.write 0
	end if%></td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<%end if%>
</TABLE>
</center>
</form>