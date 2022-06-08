function getMessage(){
	var sv=document.aspnetForm.ctl00_ContentPlaceHolder1_hf_sv.value; 
	var authName=document.aspnetForm.ctl00_ContentPlaceHolder1_hf_authName.value; 
	var ans; 

	if (sv!='') {
		ans=window.confirm('您確定要「'+sv+'」這筆'+ authName+'資料嗎？');  
	}else{
		ans=window.confirm('您確定要編修這筆'+ authName+'資料嗎？');  
	}
	if (ans==true) { 
		document.aspnetForm.ctl00$ContentPlaceHolder1$hf_confirm.value='Yes'; 
	}else { 
		document.aspnetForm.ctl00$ContentPlaceHolder1$hf_confirm.value='No';
	} 
}