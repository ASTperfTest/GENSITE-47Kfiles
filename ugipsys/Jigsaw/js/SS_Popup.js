var gcToggle = "#ffff00";
var gcBG = "#CCCCFF";

function fPopUpCalendarDlg(ctrlobj)
{
	showx = event.screenX - event.offsetX - 4 //- 210 ; // + deltaX;
	showy = event.screenY - event.offsetY + 18; // + deltaY;
	newWINwidth = 210 + 4 + 18;
	
	retval = window.showModalDialog("js/CalendarDlg.htm", "", "dialogWidth:197px; dialogHeight:210px; dialogLeft:"+showx+"px; dialogTop:"+showy+"px; status:no; directories:yes;scrollbars:no;Resizable=no; "  );
	
	
	if( retval != null ){
		//alert("已返回值！!!！");
		ctrlobj.value = retval;
	}else{
	       	//alert("canceled");
	}
}

function fPopUpCalendarDlgFormat(ctrlobj,typenum)
{
	showx = event.screenX - event.offsetX - 4 //- 210 ; // + deltaX;
	showy = event.screenY - event.offsetY + 18; // + deltaY;
	newWINwidth = 210 + 4 + 18;
	
	retval = window.showModalDialog("js/CalendarDlg.htm", "", "dialogWidth:197px; dialogHeight:235px; dialogLeft:"+showx+"px; dialogTop:"+showy+"px; status:no; directories:yes;scrollbars:no;Resizable=no; "  );
	
	
	if( retval != null ){
		//alert("已返回值！!!！");
		if (typenum == 0){
			ctrlobj.value = retval; //格式:2003-01-01
		}
		else if (typenum == 1){
			if (retval != "")
				ctrlobj.value = retval.split("-")[0] + retval.split("-")[1] + retval.split("-")[2]; //格式:20030101
			else
				ctrlobj.value = retval;
		}	
	}else{
	       	//alert("canceled");
	}
}
function fPopUpCalendarDlgFormat1(ctrlobj,typenum,formobj)
{
	showx = event.screenX - event.offsetX - 4 //- 210 ; // + deltaX;
	showy = event.screenY - event.offsetY + 18; // + deltaY;
	newWINwidth = 210 + 4 + 18;
	
	retval = window.showModalDialog("js/CalendarDlg.htm", "", "dialogWidth:197px; dialogHeight:235px; dialogLeft:"+showx+"px; dialogTop:"+showy+"px; status:no; directories:yes;scrollbars:no;Resizable=no; "  );
	
	
	if( retval != null ){
		//alert("已返回值！!!！");
		if (typenum == 0){
			ctrlobj.value = retval; //格式:2003-01-01
		}
		else if (typenum == 1){
			if (retval != "")
				ctrlobj.value = retval.split("-")[0] + retval.split("-")[1] + retval.split("-")[2]; //格式:20030101
			else
				ctrlobj.value = retval;
		}	
		
		formobj.submit();
	}else{
	       	//alert("canceled");
	}
}
