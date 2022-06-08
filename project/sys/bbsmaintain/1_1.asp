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

	sql = "select * from talk05 order by seq"
	set rs = conn.execute(sql)
	sql = "select count(*) from talk05"
	set ts = conn.execute(sql)

	page = 1
	if request("page") <> "" then	page = request("page")

	totalpage = 1
	if ts(0) > 0 then
		totalpage = ts(0) \ 10
		if ts(0) mod 10 <> 0 then	totalpage = totalpage + 1
	end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="setstyle.css">
</head>

<body>
文章分類管理

<div align="right"><input type="button" value="新增分類" onclick="javascript:location.href='1_1_addclass.asp';"></div>

<font size="2">
      共有<font color="red"><b><%= totalpage %></b></font>頁，目前在第<font color="red"><b><%= page %></b></font>頁　跳到：
      <select onchange="javascript:location.href='1_1.asp?page=' + this.value;">
<%
	n = 1
	while n <= totalpage
		response.write "<option value='" & n & "'"
		if int(n) = int(page) then	response.write " selected"
		response.write ">" & n & "</option>"
		n = n + 1
	wend
%>
      </select>頁【每頁10筆】
</font>      
      
<table cellpadding="2" cellspacing="0" border="1" bordercolordark="#FFFFFF" width="100%">
  <!--
  <tr>
    <td colspan="2">
      <input type="button" value="新增分類" onclick="javascript:location.href='1_1_addclass.asp';">
    </td>
  </tr>

  <tr>
    <td colspan="2">
      共有<%= totalpage %>頁，目前在第<%= page %>頁　跳到：
      <select onchange="javascript:location.href='1_1.asp?page=' + this.value;">
<%
	n = 1
	while n <= totalpage
		response.write "<option value='" & n & "'"
		if int(n) = int(page) then	response.write " selected"
		response.write ">" & n & "</option>"
		n = n + 1
	wend
%>
      </select>頁【每頁10筆】
    </td>
  </tr>
    -->
  <tr>
    <td width="6%" align="center" bgcolor="#eeeeee">排序</td>
    <td width="94%" bgcolor="#eeeeee">分類名稱</td>
  </tr>
<%
	i = 0
	while not rs.eof
		i = i + 1
		if i <= (page * 10) and i > (page - 1) * 10 then
%>
  <tr>
    <td width="6%" align="center"><%= i %></td>
    <td width="94%"><a href="1_1_class_fix.asp?id=<%= rs("id") %>"><%= trim(rs("name")) %></a></td>
  </tr>
<%
		end if
		rs.movenext
	wend
%>
</table>
<p><a href="javascript:history.go(-1);"><b>回上頁</b></a></p>
</body>
</html>