function controlDiv(id,st)
{
var as = "#inc_" + id;
    if ($(as).css("display")=="none"){   
        $(as).show(); 
        }
	else{
	    $(as).hide();
	}
	getxy(id,st);
}

function closeDiv(id)
{
var as = "#inc_" + id;
    if ($(as).css("display")=="none"){   
        $(as).show(); 
        }
	else{
	    $(as).hide();
	}
}

function getxy(divid,sst)
{
   divid= "inc_"+divid
   var element = sst;
   var xPos = sst.offsetLeft;
   var yPos = sst.offsetTop;
   while( sst = sst.offsetParent )
    {
        yPos += sst.offsetTop;
        xPos += sst.offsetLeft;
    }
   document.getElementById(divid).style.left = xPos + "px";
   document.getElementById(divid).style.top = yPos + 20 + "px";
}