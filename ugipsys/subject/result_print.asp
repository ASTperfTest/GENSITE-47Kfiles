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
	<td width = "85px">文章/篇</br>(含不公開)</td>
	<td width = "40px">圖片/張</td>
	<td width = "40px">附件/篇</td>
	<td>評論人數</td>
	<td>總得星數</td>
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
sql =sql & vbcrlf & "		and CuDTGeneric.fCTUPublic = 'Y'"
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

'response.write sql
set rs = conn.execute(sql)

do while not rs.eof
%>
<tr>
	<td><%=rs("CtRootName")%></td>
	<td><%=datevalue(rs("create_time"))%></td>
	<td><%=datevalue(rs("newdate"))%></td><!--最近更新日期-->
	<td><%=rs("UserName") & " / " & rs("deptName")%>&nbsp;</td>
	<td><%=rs("ClickCount") %></td><!--頁面瀏覽次數-->
	<td><%=rs("count1")%></td><!--文章-->
	<td><%=rs("count3")%></td><!--文章(含不公開)-->
	<td><%=rs("count2")%></td><!--圖片-->
	<td><%=rs("attachCount")%></td><!--附件-->
	<td><%=rs("sumP")%></td><!--評論人數-->
	<td><%=rs("sumS")%></td><!--總得星數-->
</tr>
<%
rs.movenext
loop
%>
</TABLE>
</center>
</form>