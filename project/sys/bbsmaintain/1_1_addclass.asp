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
	
	sql = "select count(*) from talk05"
	set ts = conn.execute(sql)
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="setstyle.css">
</head>

<body>
<p>�峹�����޲z �X �s�W����</p>
<table cellpadding="2" cellspacing="0" border="1" bordercolordark="#FFFFFF" width="100%">
  <tr> 
    <td colspan="2"> 
      <p><font color="#993300">�ϥλ����G</font></p>
      <ul>
        <li><font color="#993300">�����W�١G����������ܪ��W�١C</font></li>
      </ul>
    </td>
  </tr>
  <form action="1_1_addclass_act.asp" method="post" name="user">
  <tr> 
    <td>�����Ƨ�</td>
    <td><input type="text" name="seq" value="<%= ts(0) + 1 %>" size="3" maxlength="3"></td>
  </tr>
  <tr> 
    <td>�����W��</td>
    <td><input type="text" name="name" maxlength="20"></td>
  </tr>
  <tr> 
    <td>���D�W��</td>
    <td> <input type="text" name="master" maxlength="20"></td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <input type="submit" name="submit" value="�T�w">
    </td>
  </tr>
  </form>
</table>
<p><a href="javascript:history.go(-1);"><b>�^�W��</b></a></p>
</body>
</html>          