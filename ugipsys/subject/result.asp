
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
		<option value="article_nopublic" <%if request("orderby")="article_nopublic" then%>selected<%END IF%>>文章/篇(含不公開)</option>
		<option value="pic" <%if request("orderby")="pic" then%>selected<%END IF%>>圖片/張</option>
		<option value="attach" <%if request("orderby")="attach" then%>selected<%END IF%>>附件/篇</option>
		<option value="person" <%if request("orderby")="person" then%>selected<%END IF%>>評論人數</option>
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
	<th width = "85px">文章/篇</br>(含不公開)</th>
	<th width = "40px">圖片/張</th>
	<th width = "40px">附件/篇</th>
	<th>評論人數</th>
	<th>總得星數</th>
</tr>
<%

sql = ""
sql =sql & vbcrlf & "declare @startDate datetime"
sql =sql & vbcrlf & "declare @endDate datetime"

if session("DateS")<>"" then
    sql =sql & vbcrlf & "set @startDate = '"& session("DateS") &"'"
else
    sql =sql & vbcrlf & "set @startDate = '1950/01/01'"
end if
if session("DateE")<>"" then
    sql =sql & vbcrlf & "set @endDate = '"& session("DateE") &"'; set @endDate= @endDate+1"
else
    sql =sql & vbcrlf & "set @endDate = '2026-01-01'"
end if
sql =sql & vbcrlf & ""
sql =sql & vbcrlf & "select"
sql =sql & vbcrlf & "	CatTreeRoot.CtRootName"
sql =sql & vbcrlf & "	,CatTreeRoot.CtRootID"
sql =sql & vbcrlf & "	,NodeInfo.create_time"
sql =sql & vbcrlf & "	,isnull(InfoUser.UserName, '') as UserName"
sql =sql & vbcrlf & "	,isnull(dept.deptName, '') as deptName"
sql =sql & vbcrlf & "	,main.*"
sql =sql & vbcrlf & "   ,IncludeNoPubic.count3" '文章/篇(含不公開)
sql =sql & vbcrlf & "from"
sql =sql & vbcrlf & "("
sql =sql & vbcrlf & "	select "
sql =sql & vbcrlf & "		CatTreeNode.CtRootID"
sql =sql & vbcrlf & "		,COUNT(*) as count1"
sql =sql & vbcrlf & "		,COUNT(CuDTGeneric.xImgFile) as count2"
sql =sql & vbcrlf & "		,max(isnull(CuDTGeneric.dEditDate,'2000/01/01')) as newdate"
sql =sql & vbcrlf & "		,SUM(isnull(SubjectForum.CommandCount,0)) as sumP	--評論人數"
sql =sql & vbcrlf & "		,SUM(isnull(SubjectForum.GradeCount,0)) as sumS		--總得星數"	
sql =sql & vbcrlf & "		,SUM(isnull(attachments.attachCount,0)) as attachCount	--附件數"
sql =sql & vbcrlf & "		,SUM(isnull(ClickCount.dailyClick,0))  ClickCount"
sql =sql & vbcrlf & "	FROM CatTreeNode "
sql =sql & vbcrlf & "	left JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit "
sql =sql & vbcrlf & "	Left  Join SubjectForum on SubjectForum.giCUItem = CuDTGeneric.iCUItem"
sql =sql & vbcrlf & "	left  join (select iCUItem, sum(dailyClick) as dailyClick from DailyClick group by iCUItem) as ClickCount"
sql =sql & vbcrlf & "			   on ClickCount.iCUItem = CuDTGeneric.iCUItem"
sql =sql & vbcrlf & "	left  join (select xiCUItem,count(*) as attachCount from CuDTAttach group by xiCUItem) as attachments"
sql =sql & vbcrlf & "				on attachments.xiCUItem = CuDTGeneric.iCUItem"
sql =sql & vbcrlf & "	where "
sql =sql & vbcrlf & "			CatTreeNode.CtRootID in (" & session("subjects")  & ")"
sql =sql & vbcrlf & "		and CuDTGeneric.fCTUPublic = 'Y'"  '文章/篇
sql =sql & vbcrlf & "		and CuDTGeneric.xPostDate between @startDate and @endDate"
sql =sql & vbcrlf & "	group by "
sql =sql & vbcrlf & "		CatTreeNode.CtRootID"
sql =sql & vbcrlf & ")as main"
sql =sql & vbcrlf & "left join( select CatTreeNode.CtRootID ,COUNT(*) as count3"
sql =sql & vbcrlf & "	FROM CatTreeNode "
sql =sql & vbcrlf & "	left JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit"
sql =sql & vbcrlf & "	where CatTreeNode.CtRootID in (" & session("subjects")  & ")"
sql =sql & vbcrlf & "		  and CuDTGeneric.xPostDate between @startDate and @endDate"
sql =sql & vbcrlf & "	group by CatTreeNode.CtRootID"
sql =sql & vbcrlf & ") as IncludeNoPubic on IncludeNoPubic.CtRootID = main.CtRootID"
sql =sql & vbcrlf & "inner join NodeInfo on NodeInfo.CtRootID = main.CtRootID"
sql =sql & vbcrlf & "inner join CatTreeRoot on CatTreeRoot.CtRootID = main.CtRootID"
sql =sql & vbcrlf & "left join InfoUser on InfoUser.UserID=NodeInfo.owner"
sql =sql & vbcrlf & "left join dept on dept.deptId = InfoUser.deptId"
sql =sql & vbcrlf & "where CatTreeRoot.inUse = 'Y'"
select case request("orderby")
    case "click"
        sql =sql & vbcrlf & "order by ClickCount"
    case "article"
        sql =sql & vbcrlf & "order by count1"
	case "article_nopublic"
        sql =sql & vbcrlf & "order by count3"
    case "pic"
        sql =sql & vbcrlf & "order by count2"
    case "person"
        sql =sql & vbcrlf & "order by sumP"
    case "star"
        sql =sql & vbcrlf & "order by sumS"
    case "attach"
        sql =sql & vbcrlf & "order by attachCount"        
end select

' response.write sql
' response.end
set rs = conn.execute(sql)

do while not rs.eof
%>
<tr>
	<td><%=rs("CtRootName")%></td>
	<td><%=datevalue(rs("create_time"))%></td>
	<td><%=datevalue(rs("newdate"))%></td><!--最近更新日期-->
	<td><%=rs("UserName") & " / " & rs("deptName")%>&nbsp;</td>
	<td><%=rs("ClickCount") %></td><!--頁面瀏覽次數-->
	<td><a href = "result_article.asp?atype=5&id=<%=rs("CtRootID")%>"><%=rs("count1")%></a></td><!--文章/篇-->
	<td><a href = "result_article.asp?atype=1&id=<%=rs("CtRootID")%>"><%=rs("count3")%></a></td><!--文章/篇(含不公開)-->
	<td><a href = "result_article.asp?atype=2&id=<%=rs("CtRootID")%>"><%=rs("count2")%></a></td><!--圖片-->
	<td><a href = "result_article.asp?atype=3&id=<%=rs("CtRootID")%>"><%=rs("attachCount")%></a></td><!--附件-->
	<td><a href = "result_article.asp?atype=4&id=<%=rs("CtRootID")%>"><%=rs("sumP")%></a></td><!--評論人數-->
	<td><%=rs("sumS")%></td><!--總得星數-->
</tr>
<%
rs.movenext
loop
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