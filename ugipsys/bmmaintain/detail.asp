<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap="意見信箱"
HTProgFunc="意見信箱維護"
HTProgCode="BM010"
HTProgPrefix="mSession"
response.expires = 0
%>
<!-- #include virtual = "/inc/server.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
Set RSreg = Server.CreateObject("ADODB.RecordSet")

	id = request("id")
	sql = "" & _
		" select Name, Addr, Email, Tele, DATE, docid, dateline, unit, " & _
		" isnull(Context, '') context " & _
		" from Prosecute Where Id = " & replace(Id, "'", "''")
	set rs = conn.execute(sql)
	if rs.eof then
		response.write "<script language='javascript'>alert('操作錯誤！');history.go(-1);</script>"
		response.end
	end if
	
	dateline = rs("dateline")
%>
<html>
<head>
<title><%=session("mySiteName")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
</head>

<body bgcolor="#FFFFFF">
<table border="0" cellspacing="1" cellpadding="0" width="80%">
  <tr>
    <td><%=session("mySiteName")%>意見信箱</td>
  </tr>
  <tr>
    <td><hr noshade size="1" color="#000080"></td>
  </tr>
</table>
<table width=80% cellspacing="1" cellpadding="8" class=bg bgcolor=navy>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap width="10%">來信時間</td>
    <td valign="top"><font size=2><%= datevalue(rs("date")) %></font></td>
  </tr>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap>您的真實姓名</td>
    <td valign="top"><font size=2><%= trim(rs("Name")) %></font></td>
  </tr>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap>您的聯絡電話</td>
    <td valign="top"><font size=2><%= trim(rs("Tele")) %></font></td>
  </tr>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap>您的電子信箱</td>
    <td valign="top"><font size=2><a href="mailto:<%= trim(rs("EMail")) %>"><%= trim(rs("EMail")) %></a></font></td>
  </tr>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap>您的聯絡地址</td>
    <td valign="top"><font size=2><%= trim(rs("Addr")) %></font></td>
  </tr>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap>您的意見內容</td>
    <td valign="top"><font size=2><%= replace(trim(rs("Context")), vbCrLf, "<br>") %></font></td>
  </tr>

  <form action="docid_act.asp" method="post">
  <input type="hidden" name="id" value="<%= id %>">
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap>輸入文號</td>
    <td valign="top"><input type="text" name="docid" value="<%= trim(rs("docid")) %>" maxlength="20"></td>
  </tr>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap>輸入處理單位</td>
    <td valign="top">
      <select name="datafrom1">
<%
	sql_dept = "select deptid, deptname, abbrname, orgrank, kind from dept where deptid like '01%' order by deptid"
	set rs_dept = conn.execute(sql_dept)
	while not rs_dept.eof
		if trim(rs("unit")) = trim(rs_dept("deptid")) then
			response.write "<option value='" & trim(rs_dept("deptid")) & "' selected>" & trim(rs_dept("deptname")) & "</option>" & vbcrlf
		else
			response.write "<option value='" & trim(rs_dept("deptid")) & "'>" & trim(rs_dept("deptname")) & "</option>" & vbcrlf
		end if
		rs_dept.movenext
	wend
%>
      </select>
    </td>
  </tr>
<%
	if dateline <>"" then
		yy = year(dateline)
		mm = month(dateline)
		dd = day(dateline)
	end if
%>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap>處理期限</td>
    <td valign="top"><font size=2>西元<input type="text" name="yy" value="<% =yy %>" size="4" maxlength="4">年<input type="text" name="mm" value="<%= mm %>" size="2" maxlength="2">月<input type="text" name="dd" value="<%= dd %>" size="2" maxlength="2">日</font></td>
  </tr>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap>&nbsp;</td>
    <td valign="top" align="left">
      <input type="submit" name="Submit" value="確定">
      <input type="submit" name="Submit" value="刪除">
    </td>
  </tr>
  </form>
</table>
<p><a href="javascript:history.go(-1);"><b>回上頁</b></a></p>
</body>
</html>
