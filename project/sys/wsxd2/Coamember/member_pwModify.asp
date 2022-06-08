<!--#Include file = "CheckFunction.inc" -->
<script language="JavaScript">
<!--
function check(){
	var form=document.form1; 
	//if(CheckAccount(form) == false){
	//	form.account.focus();
	//	return false;
	//}
	//else 
	if(form.oldPWD.value==""){
		alert("請輸入舊密碼！"); 
		form.oldPWD.focus(); 
		return false;
	}
	else if(CheckPasswd(form) == false){
		return false;
	}
	//else if(CheckEmail(form) == false){
	//	form.email.focus();
	//	return false;
	//}
	//else if(CheckIDn(form) == false){
	//	form.idn.focus();
	//	return false;
	//}
	return true;
}  
//-->
</script>

<div class="path" >目前位置：
		<a title="首頁" href="mp.asp">首頁</a>
		&gt; 變更密碼
		</div>
		<h3>變更密碼</h3>
	<div id="Magazine">
		<div class="Event">	
		
			<div class="experts">
						<form name="form1" method="post"  class="FormA" action="sp.asp?xdURL=/site/coa/coaMember/member_pwModify_act.asp" onSubmit="return check()">
									<p>請輸入下面欄位修改您的密碼</p>
									<table  cellspacing="0" class="DataTb1">
									<tr>
										<th>請輸入您的舊密碼：</th>
										<td>
										  <input name="oldPWD" type="password" class="Text" id="oldPWD" size="30"><br/><br/>
										</td>
									</tr>
									<tr>
										<th>請輸入您的新密碼：</th>
										<td>
										  <input name="passwd" type="password" class="Text" id="passwd" size="30"><br />
										  自訂英文(區分大小寫)、數字，不含空白及@，16碼以下
										</td>
									</tr>
									<tr>
										<th>請重複ㄧ次您的新密碼：</th>
										<td>
										  <input name="passwd2" type="password" class="Text" id="passwd2" size="30">
										</td>
									</tr>
									</table>
									<Br/>
						<center>			
						<input name="Submit" type="submit" class="Button" value="確定">		
						<input type="reset" class="Button" value="重填"  >
					</center>
						</form>
			</div>


