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
	
	sql = "select content from talk03"
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
<p>������r�޲z</p>
<table width="75%" border="1">
  <form method="post" action="4_1_act.asp">
  <tr> 
    <td>�п�J�L�o��������r�G�]�H�r�����j�^</td>
  </tr>
  <tr>
    <td><textarea name="content" cols="60" rows="10"><%= content %></textarea></td>
  </tr>
  <tr> 
    <td> <input type="submit" name="submit" value="�ק�"></td>
  </tr>
  </form>  
</table>
<p><a href="javascript:history.go(-1);"><b>�^�W��</b></a></p>
</body>
</html>