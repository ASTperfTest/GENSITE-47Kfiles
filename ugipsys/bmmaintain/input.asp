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
		" select replydate, Email, DATE, Context, reply " & _
		" from Prosecute Where Id = " & replace(Id, "'", "''")
	set rs = conn.execute(sql)
	if rs.eof then
		response.write "<script language='javascript'>alert('操作錯誤！');history.go(-1);</script>"
		response.end
	end if

	if rs("replydate") <> "" then
		yy = year(rs("replydate"))
		mm = month(rs("replydate"))
		dd = day(rs("replydate"))
	end if
%>
<html>
<head>
<title><%=session("mySiteName")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<script language="JavaScript">
<!--
function send() {
	var form = document.post;
	if ( form.reply.value == "" ) {
		alert("您忘了填寫回應內容了！");
		form.reply.focus();
		return false;
	}
	if ( form.yy.value == "" || form.mm.value == "" || form.dd.value == "" ) {
		alert("您忘了填寫日期了！");
		form.yy.focus();
		return false;
	}
	return true;
}
//-->
</script>
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
  <form method="post" action="input_deal.asp" name="post" onSubmit="return send()" >
  <input type="hidden" name="id" value="<%= id %>">
  <input type="hidden" name="email" value="<%= trim(rs("email")) %>">
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap>來信時間</td>
    <td valign="top"><font size=2><%= datevalue(rs("date")) %></font></td>
  </tr>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap>副本</td>
    <td valign="top"><input type="text" name="reply_cc" size="50" maxlength="100"></td>
  </tr>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable nowrap width="10%">來信意見內容</td>
    <td valign="top"><font size=2><% if trim(rs("Context")) <> "" then response.write replace(trim(rs("Context")), vbCrLf, "<br>") %></font></td>
  </tr>
  <tr align=left bgcolor=white>
    <td valign="top" align="right" class=lightbluetable>回應內容</td>
    <td valign="top">
      <textarea name="reply" cols="65" rows="12" wrap="VIRTUAL"><%= trim(rs("reply")) %></textarea>
    </td>
  </tr>
  <!--
  <tr align=left bgcolor=white>
    <td align="right" class=lightbluetable>回應日期</td>
    <td valign="top">
      <input type="text" name="yy" size="4" maxlength=4 value="<%= yy %>">年
      <input type="text" name="mm" size="4" maxlength=2 value="<%= mm %>">月
      <input type="text" name="dd" size="4" maxlength=2 value="<%= dd %>">日
    </td>
  </tr>
  -->
  <tr align=left bgcolor=white>
    <td class=lightbluetable>&nbsp;</td>
    <td valign="top" align="left">
      <input type="submit" name="submit" value="儲存內容">
      <input type="submit" name="submit" value="寄出內容">
      <input type="reset" value="重填">
      <input type="submit" name="submit" value="刪除">
    </td>
  </tr>
  </form>
</table>
<p><a href="javascript:history.go(-1);"><b>回上頁</b></a></p>
</body>
</html>                           