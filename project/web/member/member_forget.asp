<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-TW" lang="zh-TW">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>財政部全球資訊網 - Ministry of Finance,R.O.C.- 忘記密碼</title>
<!--<link href="../css/mof.css" rel="stylesheet" type="text/css" />-->
<link href="xslgip/mof94/css/member.css" rel="stylesheet" type="text/css" />
<!--<script language="JavaScript" src="../js/mof.js" type="text/javascript"></script>-->
</head>
<body>
<div class="path"><a href="defaultp.asp">首頁</a> &amp;gt; 會員服務 &amp;gt; 忘記密碼</div>
<div class="contentword" >
	<form method="post" action="sp.asp?xdURL=member/member_forget_act.asp">
	親愛的網友，忘記您的密碼了嗎?<br/>請輸入下面欄位，系統將重新寄送新的密碼到您的電子郵件信箱中。
	<table border="0" cellpadding="3" cellspacing="0" summary="會員忘記密碼查詢欄位" class="DataTb1">
      
      <tr>
        <td colspan="2" align="right" valign="top">
        </td>
      </tr>
		  <tr>
            <th>帳　號：</th>
            <td><input name="account" type="text" class="InputBox" value="" size="30"></td>
          </tr>
          <tr>
            <th >姓　名：</th>
            <td><input name="realname" type="text" class="InputBox" value="" size="30"></td>
          </tr>
          <tr>
            <th>電子信箱：</th>
            <td><input name="email" type="text" class="InputBox" value="" size="30"></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>
		<input name="submit.x" type="submit" id="submit" value="確定"  >&nbsp;            　
		<!--<input name="cancel.x" type="submit" id="cancel" value="取消" width="44" height="17" border="0"></td>-->
		<input name="cancel.x" type="reset" id="cancel" value="取消" ></td>

          </tr>
      <tr>
        <td colspan="2" align="right" valign="top"></td>
      </tr>
</table>
</form>
</div>
</body>
</html>