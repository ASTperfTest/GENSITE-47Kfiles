<!--#Include file = "CheckFunction.inc" -->
<script language="JavaScript">
<!--
function send(){
	var form=document.form1; 
	//if(CheckAccount(form) == false){
	//	form.account.focus();
	//	return false;
	//}
	//else 
	if(form.realname.value==""){
		alert("您忘了填寫真實姓名了！"); 
		form.realname.focus(); 
		return false;
	}
	//else if(CheckPasswd(form) == false){
	//	return false;
	//}
	else 	if(CheckEmail(form) == false){
		form.email.focus();
		return false;
	}
	//else if(CheckIDn(form) == false){
	//	form.idn.focus();
	//	return false;
	//}
	
	//window.open(encodeURI("?a=sp&xdURL=/site/coa/coaMember/member_FormUpdate_act&realname=" + form.realname.value), "_blank");	
	//window.location = "?a=sp&xdURL=/site/coa/coaMember/member_FormUpdate_act&realname=" + form.realname.value + "&birthYear=" + form.birthYear.value + "&sex=" + form.sex.value + "&homeaddr=" + form.homeaddr.value + "&zip=" + form.zip.value + "&phone=" + form.zip.phone + "&mobile="+ form.mobile.value + "&fax=" + form.fax.value + "&email=" + form.email.value;
	return true;
}  
//-->
</script>
<div class="path" >目前位置：
		<a title="首頁" href="mp.asp">首頁</a>
		&gt; 查詢密碼
		</div>
		<h3>查詢密碼</h3>
	<div id="Magazine">
						<div class="Event">	

	<div class="experts">
			
			<form name="form1" method="post" action="sp.asp?xdURL=coamember/member_QueryPasswd.asp" class="FormA"  onSubmit="return send()">
			
								
				<p>查詢密碼</p>
				<p>
				請填入您之前登錄的基本資料，以便確認身分。系統會重設密碼並將密碼寄到你申請帳號時所填的 E-mail信箱。
			</p>
		<table  cellspacing="0">
          <tr>
            <th><label for="realname">＊真實姓名：</label></th>
            <td><input name="realname" id="realname"  type="text" class="Text" size="30">
              </td>
          </tr>
          <tr>
            <th><label for="email"">＊E-mail：</label></th>
            <td><input name="email" class="Text" id="email" size="30"></td>
          </tr>
        </table>
        <br/>
        <center>
			  <input name="Submit" type="submit" class="Button" value="確定">
			  <input name="Reset" type="reset" class="Button" value="清除">
			</center>
			</form>
			

    </div> 
    </div>
    </div>    