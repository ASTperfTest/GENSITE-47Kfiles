<% 
	Response.Expires = 0
	HTProgCap = "討論專區維護"
	HTProgFunc = "討論專區維護"
	HTProgCode = "BBS010"
	HTProgPrefix = "mSession" 
	response.expires = 0 
%>
<!-- #include virtual = "/inc/server.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setstyle.css">
</head>

<body>
<%= HTProgCap %>

<table width="100%" border="0" colspan="2">
  <tr>
    <td >
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr> 
    <td>
      <ul>
      <li><a href="1_1.asp">文章分類管理</a>
      <li><a href="2_1.asp">討論文章管理</a>
      <!--
      <li><a href="3_1.asp">說明文字管理</a>
      <li><a href="4_1.asp">不雅文字管理</a>
      -->
      </ul>
    </td>
  </tr>
  
</table>
<!--<p><a href="javascript:history.go(-1);"><b>回上頁</b></a></p>-->
</body>
</html>