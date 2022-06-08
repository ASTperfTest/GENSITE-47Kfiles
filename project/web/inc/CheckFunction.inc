<script language="javascript">
<!--
<!-- ----------檢查帳號 Start--------- -->
function CheckAccount(form){
	if(form.account.value==""){
		alert("您忘了填寫帳號了！"); 
		return false;
	}	
	else{
		if (form.account.value.length > 15){
			alert("您所填寫的帳號超過15碼！"); 
			return false;
		}
		var i;
		var ch;
		for (i=0;i< form.account.value.length;i++){
			ch = form.account.value.charAt(i);
			if(ch >= 'a' && ch <= 'z'){
				//return true;
			}
			else if(ch >= 'A' && ch <= 'Z'){
				//return true;
			}
			else if(ch >= '0' && ch <= '9'){
				//return true;
			}
			else if(ch == '-' || ch == '_'){
				//return true;
			}
			else{
				alert("帳號限用英文與數字，可用『-』或『_』！"); 
				return false;
			}
		}
	}
	return true;
}
<!-- ----------檢查帳號 End----------- -->

<!-- ----------檢查密碼 Start--------- -->
function CheckPasswd(form){
	if(form.passwd.value==""){
		alert("您忘了填寫密碼了！"); 
		form.passwd.focus(); 
		return false;  
	}
	else if (form.passwd.value.length > 15){
		alert("您所填寫的密碼超過15碼！"); 
		form.passwd.focus(); 
		return false;
	}
	else if(form.passwd2.value==""){
		alert("您忘了填寫再輸入密碼了！"); 
		form.passwd2.focus(); 
		return false;
	}
	else if(form.passwd2.value != form.passwd.value){
		alert("密碼與再輸入密碼不符！");
		form.passwd2.focus(); 
		return false;
	}
	else{
		var i;
		var ch;
		for (i=0;i< form.passwd.value.length;i++){		
			ch = form.passwd.value.charAt(i);		
			if(ch == '\"' || ch == ' ' || ch == '\'' || ch== '\&'){
				alert("密碼請勿包含『\"』、『'』、『&』或空白"); 
				form.account.focus();
				return false;
			}		
		}
	}
	return true;
}
<!-- ----------檢查密碼 End----------- -->

<!-- ----------檢查Email Start--------- -->
function CheckEmail(form){
	if(form.email.value==""){
		alert("您忘了填寫電子郵件了！"); 
		form.email.focus(); 
		return false;
	}
	else if(form.email.value.indexOf("@")<=-1){
		alert("您所填寫的電子郵件格式有誤！"); 
		form.email.focus(); 
		return false;
	}
	return true;
}
<!-- ----------檢查Email End----------- -->

<!-- ----------檢查身份證字號 Start--------- -->
<!-- 以下是檢查身份證字號的程式 -->
function CheckID(id) {
	var first,val,str,str1,i,v1,v2,total;
	if (id.length != 10 || (id.substr(1,1)!="1" && id.substr(1,1)!="2")) {
		alert("您填寫的身份證字號不正確1！"); 
		return false;
	}
	str="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
	val="1011121314151617341819202122352324252627282932303133";
	first=str.indexOf(id.substr(0,1).toUpperCase());
	if (first==-1) {
		alert("您填寫的身份證字號不正確2！"); 
		return false;
	}
	for (i=2; i < 9; i++) {
	if (id.substr(i,1) >= "0" && id.substr(i,1) <= "9") continue;
		alert("您填寫的身份證字號不正確3！"); 
		return false;
	}
	str1=val.substr(first*2,2);
	v1=parseInt(str1.substr(0,1));
	v2=parseInt(str1)%10;
	total=v1+v2*9+parseInt(id.substr(1,1))*8+parseInt(id.substr(2,1))*7+parseInt(id.substr(3,1))*6;
	total=total+parseInt(id.substr(4,1))*5+parseInt(id.substr(5,1))*4+parseInt(id.substr(6,1))*3;
	total=total+parseInt(id.substr(7,1))*2+parseInt(id.substr(8,1))+parseInt(id.substr(9,1));
	if ((total%10)==0) {
		return true;
	}
		alert("您填寫的身份證不正確4！"); 
		return false;
}
<!-- 以上是檢查身份證字號的程式 -->

function CheckIDn(form){
	if(form.idn.value==""){
		alert("您忘了填寫身份證字號(護照號碼)了！"); 
		return false;
	}
	else{
		if(form.foreign != null && form.foreign.value=="Y")
		{
			return true;
		}
		return CheckID(form.idn.value);
	}
	return true;
}
<!-- ----------檢查身份證字號 End----------- -->
//-->
</script>
<noscript>
  您的瀏覽器不支援script
</noscript> 