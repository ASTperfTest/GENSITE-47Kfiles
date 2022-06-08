
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

<p>Your login details.</p>

<form action="http://ngipsys.hyweb.com.tw/publicwizard/final.aspx" method="post" id="editform">
<table>
	<tr>
		<td>Email:</td>
		<td><input type="text" name="email" value="<%=request("email")%>" id="email" /></td>
	</tr>
	<tr>
		<td>Password:</td>
		<td><input type="text" name="password" value="<%=request("password")%>" id="password"/></td>
	</tr>
</table>
</form>

<script language="Javascript">
<!--

function OnBack(){
	history.go(-1);
}

function OnNext(){

	 var xml = window.external.Property("TransferManifest");
	 var files = xml.selectNodes("transfermanifest/filelist/file");

  	for (i = 0; i < files.length; i++) {
	    var postTag = xml.createNode(1, "post", "");
	    postTag.setAttribute("href", "http://ngipsys.hyweb.com.tw/publicwizard/acceptfile.aspx");
	    postTag.setAttribute("name", "myfile");
	   /*
	    var dataTag = xml.createNode(1, "formdata", "");
	    dataTag.setAttribute("name", "MAX_FILE_SIZE");
	    dataTag.text = "2000000";
	    postTag.appendChild(dataTag);
	    */
	
	    var dataTag = xml.createNode(1, "formdata", "");
	    dataTag.setAttribute("name", "action");
	    dataTag.text = "Save";
	    postTag.appendChild(dataTag);
	
	    files.item(i).appendChild(postTag);
  	}

  	var uploadTag = xml.createNode(1, "uploadinfo", "");
  	var htmluiTag = xml.createNode(1, "htmlui", "");
  	htmluiTag.text = "http://ngipsys.hyweb.com.tw/publicwizard/seeyourfiles.aspx";
  	//uploadTag.appendChild(htmluiTag);
 
  	xml.documentElement.appendChild(uploadTag);
	//alert(xml.xml);
	window.external.FinalNext();
}

function OnCancel(){
	alert('OnCancel');
}

window.external.SetWizardButtons(true, true, true);
window.external.SetHeaderText('Upload photos to GIP', 'Login with your account');
window.external.Caption = 'Upload photos to Gip';


//-->
</script>

</body>
</html>
