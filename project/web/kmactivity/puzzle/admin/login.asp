<%
If Request.Form("user_id")<>"" And Request.Form("password")<>"" Then 
	If Request.Form("user_id")="administrator" And Request.Form("password")="gss" Then
		session("login")="ok"
		Response.Redirect "main.asp"
	Else
		session("errnumber")=1
	  	session("msg")="帳號或密碼錯誤！"	
	End If
End If

Sub Message()
  If session("errnumber")=0 then
    Response.Write "<center>"&session("msg")&"</center>"
  Else
    Response.Write "<script Language='JavaScript'>alert('"&session("msg")&"')</script>"
  End If
  session("msg")=""
  session("errnumber")=0
End Sub
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=big5" />
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
  <title>圖庫管理</title>
  <link href="login.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
</script>
</head>
<body class="body" OnLoad="Window_OnLoad()">
<div id="container">
  <p>&nbsp;</p>
  <div id="login">
  <h1>圖庫管理</h1>
  <form name="form" method="POST" action="">
  <table border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <th width="60" align="center"> 帳 號</th>
      <td width="180" valign="middle"><input name="user_id" type="text" class="keyin" size="20" maxlength="20"></td>
    </tr>
    <tr>
      <th align="center"> 密 碼</th>
      <td valign="middle"><input name="password" type="password" class="keyin" size="21" maxlength="10"></td>
    </tr>
    <tr>
      <td colspan="2" align="center"><input type="submit" value=" 登 入 " name="action" style="font-family: 新細明體; font-size: 9pt"></td>
    </tr>
  </table>
  </form> 
  </div>
  <div id="footer">
    <p>&nbsp;</p>
  </div>
</div>
<%Message()%>
</body>
</html>
  
<script language="JavaScript"><!--
  function Window_OnLoad(){
    document.form.user_id.focus();
  }
--></script>