<%
	Response.Expires = 0
	HTProgCap = "�Q�ױM��"
	HTProgFunc = "�Q�ױM�Ϻ��@"
	HTProgCode = "BBS010"
	HTProgPrefix = "mSession"
	response.expires = 0
%>
<!-- #include virtual = "/inc/server.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")


	bid = trim(request("bid"))
	if bid = "" then
		response.redirect "index.asp"
		response.end
	end if

	sql = "select name from talk05 where id = " & pkstr(bid, "")
	set rs = conn.execute(sql)
	if rs.eof then
		response.redirect "index.asp"
		response.end
	end if
	session("bbs_name") = trim(rs("name"))


	sql = "select count(*) from article where bid = " & pkstr(bid, "")
	set ts = conn.execute(sql)

	pagecount = 10
	page = 1
	if request("page") <> "" then	page = request("page")

	totalpage = 1
	if ts(0) > 0 then
		totalpage = ts(0) \ pagecount
		if ts(0) mod pagecount <> 0 then	totalpage = totalpage + 1
	end if
%>
<html>
<head>
<title>�Q�ױM��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<link rel="stylesheet" href="setstyle.css">
</head>

<body>
<p>�Q�פ峹�޲z - <%= session("bbs_name") %></p>

<font size="2">
      �@��<font color="red"><b><%= totalpage %></b></font>���A�ثe�b��<font color="red"><b><%= page %></b></font>���@����G
      <select onchange="javascript:location.href='articlelist.asp?bid=<%= bid %>&page=' + this.value;">
<%
	n = 1
	while n <= totalpage
		response.write "<option value='" & n & "'"
		if int(n) = int(page) then	response.write " selected"
		response.write ">" & n & "</option>"
		n = n + 1
	wend
%>
      </select>���i�C��<%= pagecount %>���j	
</font>

<table width="100%" border="1">
  <tr>
    <td width="5%" nowrap>�s��</td>
    <td nowrap>���D</td>
    <td width="20%" nowrap>�@��</td>
    <td width="20%" nowrap>�i�K��</td>
    <!--<td align="right" width="10%" nowrap>�s���H��</td>-->
  </tr>
  <!--
  <tr>
    <td colspan="5"><hr></td>
  </tr>
  -->
<%
	sql = "" & _
		" select nickname, date, id, mid, sid, bid, title, cnt, " & _
		" isnull(isethics, 0) isethics " & _
		" from article where bid = " & pkstr(bid, "") & _
		" order by mid desc, sid "
	set rs = conn.execute(sql)
	i = 0
	while not rs.eof
		i = i + 1
		if i > (page - 1) * pagecount and i <= page * pagecount Then
			if int("0" & pre_mid) <> int(rs("mid")) then
				if color_str = "" then
					color_str = " bgcolor=#DDDDDD"
				else
					color_str = ""
				end if
				pre_mid = rs("mid")
			end if
%>
  <tr<%= color_str %>>
    <td align="center"><%= i %></td>
    <td><a href=Article.asp?MId=<%= rs("MID") %>&SId=<%= rs("SID") %>&bId=<%= BId %>><%= trim(rs("Title")) %></a><% if rs("isethics") = "1" then	response.write "<img src=images/ethics_logo.gif border=0>" %></td>
    <td nowrap><%= trim(rs("NickName")) %></td>
    <td nowrap><%= rs("date") %></td>
    <!--<td align="right" nowrap><%= rs("cnt") %>��</td>-->
  </tr>
<%
		end if
		rs.movenext
	wend
%>
<!--
  <tr>
    <td colspan="5"><hr></td>
  </tr>
  -->
  <!--
  <tr>
    <td colspan="5">
      �@��<%= totalpage %>���A�ثe�b��<%= page %>���@����G
      <select onchange="javascript:location.href='articlelist.asp?bid=<%= bid %>&page=' + this.value;">
<%
	n = 1
	while n <= totalpage
		response.write "<option value='" & n & "'"
		if int(n) = int(page) then	response.write " selected"
		response.write ">" & n & "</option>"
		n = n + 1
	wend
%>
      </select>���i�C��<%= pagecount %>���j
    </td>
  </tr>
-->  
</table>
<p><a href="javascript:history.go(-1);"><b>�^�W��</b></a></p>
</body>
</html>
