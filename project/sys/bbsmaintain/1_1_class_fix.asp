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
	
	id = trim(request("id"))
	if id = "" then
		response.write "<script language='javascript'>history.go(-1);</script>"
		response.end
	end if
	
	
	sql = "select * from talk05 where id = " & pkstr(id, "")
	set rs = conn.execute(sql)
	if rs.eof then
		response.write "<script language='javascript'>history.go(-1);</script>"
		response.end
	end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="setstyle.css">
</head>

<body>
<p>�峹�����޲z �X �����ק�</p>
<table cellpadding="2" cellspacing="0" border="1" bordercolordark="#FFFFFF" width="770">
  <tr> 
    <td colspan="2"> 
      <p><font color="#993300">�ϥλ����G</font></p>
      <ul>
        <li><font color="#993300">�����W�١G�п�J�����W�١C</font></li>
      </ul>
    </td>
  </tr>
  <form action="1_1_class_fix_act.asp" method="post" name="user">
  <input type="hidden" name="id" value="<%= id %>">
  <tr> 
    <td>�����Ƨ�</td>
    <td><input type="text" name="seq" size="3" maxlength="3" value="<%= rs("seq") %>"></td>
  </tr>
  <tr> 
    <td>�����W��</td>
    <td><input type="text" name="name" value="<%= trim(rs("name")) %>" maxlength="20"></td>
  </tr>
  <tr> 
    <td>���D�W��</td>
    <td><input type="text" name="master" value="<% =trim(rs("master")) %>" maxlength="20"></td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <input type="submit" name="submit" value="�ק�"> 
      <input type="submit" name="submit" value="�R��">
    </td>
  </tr>
  </form>
</table>
<p><a href="javascript:history.go(-1);"><b>�^�W��</b></a></p>
</body>
</html>