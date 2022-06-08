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
	
	
	bid = trim(request("bid"))
	mid2 = trim(request("mid"))
	sid = trim(request("sid"))
	if bid = "" or mid2 = "" or sid = "" then
		response.redirect "index.asp"
		response.end
	end if
	
	sql = "" & _
		" select nickname, date, title, message " & _
		" from article where mid = " & pkstr(mid2, "") & _
		" and sid = " & pkstr(sid, "") & _
		" and bid = " & pkstr(bid, "")
	set rs = conn.execute(sql)
	if rs.eof then
		response.redirect "index.asp"
		response.end
	end if
%>
<html>
<head>
<title>討論專區</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<link rel="stylesheet" href="setstyle.css">
</head>

<body>
<p>討論文章管理 - <%= session("bbs_name") %></p>
<table width="100%" border="0"">
  <form action="article_fix.asp" method="post">
  <input type="hidden" name="sid" value="<%= sid %>">
  <input type="hidden" name="mid" value="<%= mid2 %>">
  <input type="hidden" name="bid" value="<%= bid %>">
  <tr>
    <td align="right" width="15%" nowrap>作者：</td>
    <td><%= trim(rs("NickName")) %></td>
  </tr>
  <tr>
    <td align="right">標題：</td>
    <td><input type="text" name="title" size="45" maxlength="30" value="<%= trim(rs("Title")) %>"></td>
  </tr>
  <tr>
    <td align="right">內容：</td>
    <td><textarea name="message" wrap="VIRTUAL" cols="50" rows="10"><%= trim(rs("message")) %></textarea></td>
  </tr>
  <tr>
    <td align="right" nowrap>張貼時間：</td>
    <td><%= rs("date") %></td>
  </tr>
  <tr>
    <td></td>
    <td>
      <input type="submit" name="submit" value="修改">
      <input type="submit" name="submit" value="刪除">
    </td>
  </tr>
</table>
<p><a href="javascript:history.go(-1);"><b>回上頁</b></a></p>
</body>
</html>
