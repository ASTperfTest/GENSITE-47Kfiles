﻿ 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>農委會農業知識入口網維護平台</title>
<link href="xslgip/intra1/css/gip.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>
<body onLoad="Javascript:windowOnLoad();">


<div class="doclink">
<span class="forget">
<img src="xslgip/intra1/images/forgetpassword.gif" width="28" height="26" align="absmiddle" />
<!--a href="#" >忘記密碼</a-->
<a href="mailto:km@mail.coa.gov.tw"">聯絡管理員</a>
</span>
<span class="doc">
<img src="xslgip/intra1/images/relateddoc.gif" width="28" height="26" align="absmiddle" />
<a href="doorlist.htm">回網站選擇頁</a>
</span>
</div>

								
<form method="POST" action="checkLogin.asp" name="f1" target="_self" id="f1">
<table class="login">
<tr><td>維護網站別：</td><td style="color:red"><%=session("mySiteName")%></td></tr>
<tr>
<td>使用者名稱：
</td>
<td><input name="tfx_userName" type="text" class="inputtext" size="22" /></td>
</tr>
<tr>
<td>使用者密碼：
</td>
<td><input name="tfx_passWord" type="password" class="inputtext" size="22" /></td>
</tr>
<tr>
<td colspan="2"><input name="Submitb" type="button" class="inputbutton" value="送出" OnClick="Javascript:xsubmit();" />
<input name="reset" type="reset" class="inputbutton" value="重設" /></td>
</tr>
</table>
</form>					
			
<p class="copyright">Hyweb Government Information Potal<br />
&copy; 2007 Copyright by Hyweb Technology Co. Inc. All right reserved.</p>
</body>
</html>
<script TYPE="text/javascript">
<!--
function fun1(evnt) {   
	evnt = evnt || window.event;
	if ((evnt.which || evnt.keyCode) == 13) { xsubmit(); }
}
document.onkeypress = fun1;

function IDQ() {
        window.open("ReID.asp","忘記密碼","width=320,height=180");
}
function windowOnLoad() {
	document.f1.tfx_userName.focus();
}
function xsubmit() {
  if (document.f1.tfx_userName.value=="")
  {
     	alert("帳號 未填！");
	document.f1.tfx_userName.focus();
     	return false;
  } 
  if (document.f1.tfx_passWord.value=="")
  {
     	alert("密碼 未填！");
	document.f1.tfx_passWord.focus();
     	return false;
  } 
  document.f1.submit();
}
-->
</script>
<script language=vbs>
sub document_onkeydown()
  if window.event.keyCode = 13 then 
	if f1.tfx_userName.value = "" then
		msgBox "帳號 未填！"
    		f1.tfx_userName.focus
		window.event.returnvalue=false
		exit sub
	end if
	if f1.tfx_passWord.value = "" then
		msgBox "密碼 未填！"
    		f1.tfx_passWord.focus
		window.event.returnvalue=false
		exit sub
	end if
	f1.submit()
  End If
end sub
</script>
