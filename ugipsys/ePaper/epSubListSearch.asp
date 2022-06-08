<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="電子報管理"
HTProgFunc="查詢電子報"
HTProgCode="GW1M51"
HTProgPrefix="epSub" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
	<tr>
		<td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
			<A href="epSubList.asp?epTreeID=<%=Request("epTreeID")%>" title="回訂閱清單">回訂閱清單</A> 
		</td>
	</tr>
	<tr>
		<td width="100%" colspan="2">
			<hr noshade size="1" color="#000080">
		</td>
	</tr>
	<tr>
		<td class="Formtext" colspan="2" height="15"></td>
	</tr>  
</table>
<br>
<table align="center">
<form method=post action="epSubList.asp?epTreeID=<%=Request("epTreeID")%>" name="form1">
	<tr> 
		<td width="30%" align="center"> 
			會員ID：
		</td>
		<td width="60%"> 
			<input type="text" name="account" >
		</td>
	</tr>
	<tr> 
		<td width="30%" align="center"> 
			會員姓名：
		</td>
		<td width="60%"> 
			<input type="text" name="realname" >
		</td>
	</tr>
	<tr> 
		<td width="30%" align="center">  
			會員E-mail：
		</td>
		<td width="60%"> 
			<input type="text" name="eMail" >
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center">
			<input type="submit" name="Submit" value="查詢">
			<input type="Reset" name="reset" value="重填">
		</td>
	</tr>
</form>
</table>
