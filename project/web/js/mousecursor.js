var sTop = document.documentElement.scrollTop || document.body.scrollTop || 0;
		var sLeft = document.documentElement.scrollLeft || document.body.scrollLeft || 0;
		var xM=0;
		var yM=0;
		var png="";
		<xsl:if test="//MpStyle/Article/Caption='1'">png="springpng";</xsl:if>
		<xsl:if test="//MpStyle/Article/Caption='2'">png="summerpng";</xsl:if>
		<xsl:if test="//MpStyle/Article/Caption='3'">png="autumnpng";</xsl:if>
		<xsl:if test="//MpStyle/Article/Caption='4'">png="winterpng";</xsl:if>
		<xsl:if test="//MpStyle/Article/Caption='5'">png="autumnpng";</xsl:if>
		<xsl:if test="//MpStyle/Article/Caption='6'">png="summerpng";</xsl:if>
		<xsl:if test="//MpStyle/Article/Caption='7'">png="winterpng";</xsl:if>
		if(png.length!=0)
		{
		    document.getElementById("cursor-img").src = "/xslgip/<xsl:value-of select="myStyle"/>/images/"+((getInternetExplorerVersion()==6.0)?png+"_ie6.gif":png+".png");
		    document.getElementById("cursor-img").style.display = '';
		}
		var mousex=0;
		var mousey=0;
		function moveCursor(x, y, sT, sL)
		{
			documentBody = GetDocumentBody();
			cursorInnerWidth = documentBody.clientWidth;  
			cursorInnerHeight = documentBody.clientHeight;
			if((cursorInnerWidth-60)<=x){
				x = x -80;}
			if((cursorInnerHeight -50)<=y ){
				y = y -50;}
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
		    sTop = document.documentElement.scrollTop || document.body.scrollTop || 0;
		    sLeft = document.documentElement.scrollLeft || document.body.scrollLeft || 0;
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