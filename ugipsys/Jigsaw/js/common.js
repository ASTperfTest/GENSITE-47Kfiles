
  function fPopUpCalendarDlgFormat(ctrlobj,ctrlobjC,typenum)
                {
                showx = event.screenX - event.offsetX - 4 //- 210 ; // + deltaX;
                showy = event.screenY - event.offsetY + 18; // + deltaY;
                newWINwidth = 210 + 4 + 18;

                retval = window.showModalDialog("js/calendar/calendar.htm", "", "dialogWidth:197px; dialogHeight:210px; dialogLeft:"+showx+"px; dialogTop:"+showy+"px; status:no; directories:yes;scrollbars:no;Resizable=no; "  );

                if( retval != null ){
                //
                if (typenum == 0){
                ctrlobj.value = retval.split(',')[0]; //2007-01-01
                ctrlobjC.value = retval.split(',')[1]; //96-01-01
                }
                else if (typenum == 1){
                if (retval != "")
                ctrlobj.value = retval.split("-")[0] + retval.split("-")[1] + retval.split("-")[2]; //
              //  ctrlobjC.value = retval.split("-")[0] + retval.split("-")[1] + retval.split("-")[2]; //              
                else
                ctrlobj.value = retval;
                //ctrlobjC.value = retval;
                }
                }else{
                //alert("canceled");
                }
                }
                
 
