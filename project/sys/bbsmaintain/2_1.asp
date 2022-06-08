<% 
	Response.Expires = 0
	HTProgCap = "討論專區"
	HTProgFunc = "討論專區維護"
	HTProgCode = "BBS010"
	HTProgPrefix = "mSession" 
	response.expires = 0 
%>
<!-- #include virtual = "/inc/server.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<html>
<head>
<title>討論專區</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<link rel="stylesheet" href="setstyle.css">
</head>

<body>
<p>討論文章管理</p>
<table width="75%" border="0">
  <tr>
    <td nowrap>文章分類</td>
    <td align="right">文章篇數</td>
  </tr>
  <tr>
    <td colspan="2"><hr></td>
  </tr>
<%
	sql = ""
	sql = sql &	" select a.id, a.name, count(b.bid) cnt " 
	sql = sql &	" from talk05 a left join article b on a.id = b.bid " 
	'sql = sql &	" where a.master = '" & session("userid") & "' " 
	sql = sql &	" group by a.id, a.name having count(b.bid) > 0 "
	'sql = "select * from Article"
	set rs = conn.execute(sql)
	'response.write sql
	'response.end
	while not rs.eof
%>
  <tr>
    <td nowrap><a href="articlelist.asp?bid=<%= rs("id") %>"><%= trim(rs("name")) %></a></td>
    <td align="right"><%= rs("cnt") %>篇</td>
  </tr>
<%
		rs.movenext
	wend
%>
</table>
<p><a href="javascript:history.go(-1);"><b>回上頁</b></a></p>
</body>
</html>
