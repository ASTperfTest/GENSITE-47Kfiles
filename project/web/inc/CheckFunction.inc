<script language="javascript">
<!--
<!-- ----------�ˬd�b�� Start--------- -->
function CheckAccount(form){
	if(form.account.value==""){
		alert("�z�ѤF��g�b���F�I"); 
		return false;
	}	
	else{
		if (form.account.value.length > 15){
			alert("�z�Ҷ�g���b���W�L15�X�I"); 
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
				alert("�b�����έ^��P�Ʀr�A�i�Ρy-�z�Ρy_�z�I"); 
				return false;
			}
		}
	}
	return true;
}
<!-- ----------�ˬd�b�� End----------- -->

<!-- ----------�ˬd�K�X Start--------- -->
function CheckPasswd(form){
	if(form.passwd.value==""){
		alert("�z�ѤF��g�K�X�F�I"); 
		form.passwd.focus(); 
		return false;  
	}
	else if (form.passwd.value.length > 15){
		alert("�z�Ҷ�g���K�X�W�L15�X�I"); 
		form.passwd.focus(); 
		return false;
	}
	else if(form.passwd2.value==""){
		alert("�z�ѤF��g�A��J�K�X�F�I"); 
		form.passwd2.focus(); 
		return false;
	}
	else if(form.passwd2.value != form.passwd.value){
		alert("�K�X�P�A��J�K�X���šI");
		form.passwd2.focus(); 
		return false;
	}
	else{
		var i;
		var ch;
		for (i=0;i< form.passwd.value.length;i++){		
			ch = form.passwd.value.charAt(i);		
			if(ch == '\"' || ch == ' ' || ch == '\'' || ch== '\&'){
				alert("�K�X�Фť]�t�y\"�z�B�y'�z�B�y&�z�Ϊť�"); 
				form.account.focus();
				return false;
			}		
		}
	}
	return true;
}
<!-- ----------�ˬd�K�X End----------- -->

<!-- ----------�ˬdEmail Start--------- -->
function CheckEmail(form){
	if(form.email.value==""){
		alert("�z�ѤF��g�q�l�l��F�I"); 
		form.email.focus(); 
		return false;
	}
	else if(form.email.value.indexOf("@")<=-1){
		alert("�z�Ҷ�g���q�l�l��榡���~�I"); 
		form.email.focus(); 
		return false;
	}
	return true;
}
<!-- ----------�ˬdEmail End----------- -->

<!-- ----------�ˬd�����Ҧr�� Start--------- -->
<!-- �H�U�O�ˬd�����Ҧr�����{�� -->
function CheckID(id) {
	var first,val,str,str1,i,v1,v2,total;
	if (id.length != 10 || (id.substr(1,1)!="1" && id.substr(1,1)!="2")) {
		alert("�z��g�������Ҧr�������T1�I"); 
		return false;
	}
	str="ABCDEFGHIJKLMNOPQRSTUVWXYZ";
	val="1011121314151617341819202122352324252627282932303133";
	first=str.indexOf(id.substr(0,1).toUpperCase());
	if (first==-1) {
		alert("�z��g�������Ҧr�������T2�I"); 
		return false;
	}
	for (i=2; i < 9; i++) {
	if (id.substr(i,1) >= "0" && id.substr(i,1) <= "9") continue;
		alert("�z��g�������Ҧr�������T3�I"); 
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
		alert("�z��g�������Ҥ����T4�I"); 
		return false;
}
<!-- �H�W�O�ˬd�����Ҧr�����{�� -->

function CheckIDn(form){
	if(form.idn.value==""){
		alert("�z�ѤF��g�����Ҧr��(�@�Ӹ��X)�F�I"); 
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
<!-- ----------�ˬd�����Ҧr�� End----------- -->
//-->
</script>
<noscript>
  �z���s�������䴩script
</noscript> 