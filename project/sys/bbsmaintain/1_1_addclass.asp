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
<p>文章分類管理 — 新增分類</p>
<table cellpadding="2" cellspacing="0" border="1" bordercolordark="#FFFFFF" width="100%">
  <tr> 
    <td colspan="2"> 
      <p><font color="#993300">使用說明：</font></p>
      <ul>
        <li><font color="#993300">分類名稱：分類項目顯示的名稱。</font></li>
      </ul>
    </td>
  </tr>
  <form action="1_1_addclass_act.asp" method="post" name="user">
  <tr> 
    <td>分類排序</td>
    <td><input type="text" name="seq" value="<%= ts(0) + 1 %>" size="3" maxlength="3"></td>
  </tr>
  <tr> 
    <td>分類名稱</td>
    <td><input type="text" name="name" maxlength="20"></td>
  </tr>
  <tr> 
    <td>版主名稱</td>
    <td> <input type="text" name="master" maxlength="20"></td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <input type="submit" name="submit" value="確定">
    </td>
  </tr>
  </form>
</table>
<p><a href="javascript:history.go(-1);"><b>回上頁</b></a></p>
</body>
</html>          