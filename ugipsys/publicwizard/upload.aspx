
<html>
<head>
<title>XP Publish Wizard</title>
<style>


body {
	font: message-box;
	margin: 20px;
	padding: 0px;
	color: buttontext;
	background-color: buttonface;
}

td {
	font: message-box;
}

form {
	padding: 0px;
	margin: 0px;
}

.BuddyIconH {
	margin-right: 20px;
	margin-top: 0px;
	margin-bottom: 0px;
}

h1 {
	font: normal 32px;
	margin-top: 0px;
	margin-left: 0px;
}

h4 {
	font: normal 16px;
	margin-top: 0px;
	margin-left: 0px;
}

img {
	border: 1px solid #000000;
}


</style>
</head>
<body onContextMenu="event.returnValue=false;">

<p>Please enter your login details.</p>

<form action="http://ngipsys.hyweb.com.tw/publicwizard/final.aspx" method="post" id="editform">
<table>
	<tr>
		<td>Email:</td>
		<td><input type="text" name="email" value="aaa" id="email" onchange="editme();" onkeyup="editme();" /></td>
	</tr>
	<tr>
		<td>Password:</td>
		<td><input type="password" name="password" value="bbb" id="password" onchange="editme();" onkeyup="editme();" /></td>
	</tr>
</table>
</form>

<script language="Javascript">
<!--


function editme(){
	var email = document.getElementById('email').value;
	var password = document.getElementById('password').value;

	var email_ok = (email != '') && (email != undefined);
	var password_ok = (password != '') && (password != undefined);

	var can_submit = email_ok && password_ok;

	window.external.SetWizardButtons(true, can_submit, false);
}

function OnBack(){
	window.external.FinalBack();
}

function OnNext(){
	document.getElementById('editform').submit();
}

function OnCancel(){
	alert('OnCancel');
}

window.external.SetWizardButtons(true, true, false);
window.external.SetHeaderText('Upload photos to GIP', 'Login with your account');
window.external.Caption = 'Upload photos to Gip';


//-->
</script>

</body>
</html>
