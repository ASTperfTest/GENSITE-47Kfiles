<%@ CodePage = 65001 %><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>無標題文件</title>
		<style type="text/css">
			<!--
body {
	background-color: #161713;
	margin-left: 0%;
	margin-top: 100px;
	margin-right: 0%;
	margin-bottom: 100px;
}
--></style>
		<link href="xslgip/intra1/css/gip.css" rel="stylesheet" type="text/css" />
		<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
		</script>
	</head>
	<body>
		<div align="center">
			<table width="642" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="642" height="194" align="right" valign="top" background="xslgip/intra1/images/login_bk_top.jpg">
						<span class="doclink">
							<img src="xslgip/intra1/images/forgetpassword.gif" width="28" height="26" align="absmiddle" />
							<a href="#" onClick="VBS:IDQ">忘記密碼</a></span>
						<span class="doclink">
							<img src="xslgip/intra1/images/relateddoc.gif" width="28" height="26" align="absmiddle" />
							<a href="javascript:MM_openBrWindow('document/list.asp','document','scrollbars=yes,resizable=yes,width=488,height=250')">
								相關操作手冊&nbsp;&nbsp;</a></span></td>
				</tr>
				<tr>
					<td height="100" align="center" valign="middle" bgcolor="#D7F1A1"><strong class="customername"><%=session("mySiteName")%>, 
							請輸入使用者資訊</strong><br />
						<form method="POST" action="checkLogin_rdecthinktank.asp" name="f1" target="_self" id="f1">
							<table width="250" border="0" cellspacing="1" cellpadding="1">
								<tr>
									<td width="75" valign="middle"><img src="xslgip/intra1/images/usename.gif" width="70" height="17" />
									</td>
									<td align="left"><input name="tfx_userName" type="text" class="inputtext" size="16" /></td>
								</tr>
								<tr>
									<td width="75" valign="middle"><img src="xslgip/intra1/images/password.gif" width="70" height="17" />
									</td>
									<td align="left"><input name="tfx_passWord" type="password" class="inputtext" size="16" /></td>
								</tr>
								<tr align="center">
									<td colspan="2" valign="middle"><input name="Submitb" type="button" class="inputbutton" value="送出" OnClick="xsubmit()" />
										&nbsp;&nbsp; <input name="reset" type="reset" class="inputbutton" value="重設" /></td>
								</tr>
							</table>
						</form>
					</td>
				</tr>
				<tr>
					<td width="642" height="133" align="center" valign="bottom" background="xslgip/intra1/images/login_bk_down.jpg"><span class="copyright">HYWEB
      GOVERMENT INFORMATION POTAL<br />
&copy; 2004 Copyright by Hyweb Technology Co. Inc. All right reserved.</span><br />
						<br />
					</td>
				</tr>
			</table>
		</div>
	</body>
</html>
<script language="vbs">
sub IDQ
        window.open "ReID.asp","忘記密碼","width=300px,height=150px"
end sub

sub window_onLoad
    f1.tfx_userName.focus
end sub

sub xsubmit
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
end sub

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
