var sTop = document.documentElement.scrollTop  || document.body.scrollTop || 0;
		var sLeft = document.documentElement.scrollLeft  || document.body.scrollLeft || 0;
		var xM=0;
		var yM=0;
		  
		if(Cursorpng!=null && Cursorpng.length!=0)
		{
		    document.getElementById("cursor-img").src = "images/"+((getInternetExplorerVersion()==6.0)?Cursorpng+"_ie6.gif":Cursorpng+".png");
		    document.getElementById("cursor-img").style.display = '';
		}
		var mousex=0;
		var mousey=0;
		function moveCursor(x, y, sT, sL)
		{
			if( document.body.scrollTop ){
				documentBody = document.body;
			  }else{
				documentBody = document.documentElement;
			  }
			cursorInnerWidth = documentBody.clientWidth;  
			cursorInnerHeight = documentBody.clientHeight;
			if((cursorInnerWidth-55)<=x)
				x = cursorInnerWidth -55;
			if((cursorInnerHeight -45)<=y )
				y = cursorInnerHeight -45;
		    document.getElementById("cursor").style.left = (x + sL + 20) + "px";
		    document.getElementById("cursor").style.top =  (y + sT) + "px";
			mousex = x ;
			mousey = y ;
		}
		function mousemove(evt)
		{
		    sTop = document.documentElement.scrollTop || document.body.scrollTop || 0;
		    sLeft = document.documentElement.scrollLeft || document.body.scrollLeft || 0;
		    evt=evt || window.event;
		    xM = (evt.clientX);
		    yM = (evt.clientY);
		    moveCursor(xM, yM, sTop, sLeft);
        document.getElementById("cursor").style.display = '';
		}
		function scrollmove()
		{
		    sTop = document.documentElement.scrollTop || document.body.scrollTop ||  0;
		    sLeft = document.documentElement.scrollLeft || document.body.scrollLeft ||  0;
		    //document.getElementById("cursor").style.display = 'none';
		    moveCursor(mousex, mousey, sTop, sLeft);
		    //document.getElementById("cursor").style.display = '';
		}
		function getInternetExplorerVersion()
		{
		   var rv = -1; // Return value assumes failure.
		   if (navigator.appName == 'Microsoft Internet Explorer')
		   {
		      var ua = navigator.userAgent;
		      var re  = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
		      if (re.exec(ua) != null)
			 rv = parseFloat( RegExp.$1 );
		   }
		   return rv;
		}
		function GetDocumentBody()
		{
			if( document.body.scrollTop ){
				return document.body;
			  }else{
				return document.documentElement;
			  }
		}
		window.document.onmousemove= mousemove;
		window.onscroll = scrollmove;