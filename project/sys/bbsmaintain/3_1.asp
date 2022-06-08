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
	
	sql = "select content from talk02"
	set rs = conn.execute(sql)
	if not rs.eof then	content = trim(rs("content"))
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="setstyle.css">
</head>

<body>
<p>說明文字管理</p> 
<table width="75%" border="1">
  <form method="post" action="3_1_act.asp">
  <tr> 
    <td valign="top" width="15%" nowrap>原有說明文字：</td>
    <td><%= replace(content, vbcrlf, "<br>") %></td>
  </tr>
  <tr> 
    <td valign="top" nowrap>修改說明文字： </td>
    <td><textarea name="content" cols="60" rows="10" wrap="VIRTUAL"></textarea></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td> 
      <input type="submit" name="submit" value="修改">
    </td>
  </tr>
  </form>
</table>
<p><a href="javascript:history.go(-1);"><b>回上頁</b></a></p>
</body>
</html>