var flag;
var selected = new Array();
var arraylen = 0;

function addloc(locs,mylocs){
  for(var x=0;x<locs.length;x++){
    var opt = locs.options[x];
    if (opt.selected){
      flag = true;
      for (var y=0;y<mylocs.length;y++){
        var myopt = mylocs.options[y];
        if (myopt.value == opt.value){  
          flag = false;
        }
      } 
      if(flag){
        mylocs.options[mylocs.options.length] = new Option(opt.text, opt.value, 0, 0); 
        if(window.top.selected!=null ){
            var s = new Array(opt.text,opt.value);
            window.top.selected[window.top.arraylen]=s;
            window.top.arraylen++;
        }
      }
    }
  }
}

function copyoldIDV(mylocs,olds) {
  for (var x = 0; x < olds.length; x++ ) {
    var old = olds[x];		
    if ( old[0] == null || old[0] == "$valuestr" || old[0] == "" )
      return;			
    if ( arraylen < olds.length ) {				
			var s = new Array(old[1],old[0]);
      selected[arraylen] = s;
      arraylen++; 
    }
  }
}

//當頁面裝載時重新裝載所有已經選擇的選項
function loadLoc(mylocs){
    if(selected!=null && selected.length!=0){
		for (var y=0;y < selected.length;y++){
	   		 var s = selected[y];	
	   		 if(s!=null)					
        	    mylocs.options[mylocs.options.length] = new Option(s[1], s[0], 0, 0);
     	 } 
     }
}

function delloc(locs,mylocs){
  for(var x=mylocs.length-1;x>=0;x--){
    var opt = mylocs.options[x];
    if (opt.selected){
      mylocs.options[x] = null;
      if(window.top.selected!=null){
        for (var y=0;y<window.top.selected.length;y++){
	   		 var s = window.top.selected[y];
	   		 if(s!=null && s[0]==opt.text && s[1]==opt.value){
        	   window.top.selected[y] = null; 
        	    break;
        	 }
     	 }
      }
    }
  }
}

 function copyold(locs,mylocs,olds){
 	for(i=0;i<locs.length;i++){
 		var lopt = locs.options[i];
 		if(isold(lopt.value,olds)){
 			mylocs.options[mylocs.options.length] = new Option(lopt.text, lopt.value, 0, 0);
 		}
 	}
 }
 function isold(value,olds){
 	for(j=0;j<olds.length;j++){
 		if(value==olds[j])
 			return true;
 	}
 	return false;
 }