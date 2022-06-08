<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
xmlns:hyweb="urn:gip-hyweb-com" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
<xsl:include href="member.xsl"/>
<xsl:include href="../hwXslFunc.xsl"/>
<xsl:variable name="title">農業知識入口網</xsl:variable>

<!-- Header／開始 -->
<!--var fornothing = 00000; 是因為知識家urchin.js下一個javascript區塊不會運作而放置的 -->
<xsl:template match="hpMain" mode="Header">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title><xsl:value-of select="$title"/> －小知識串成的大力量－<xsl:value-of select="//xPath/UnitName"/>/<xsl:value-of select="//MainArticle/Caption"/></title>
    <script language="javascript" type="text/javascript" src="/js/jquery.js">/**/</script>
    <link rel="stylesheet" type="text/css">
		<xsl:attribute name="href">/xslgip/<xsl:value-of select="myStyle"/>/css/4seasons.css</xsl:attribute>
	</link>
	<link rel="stylesheet" type="text/css">
		<xsl:attribute name="href">/xslgip/<xsl:value-of select="myStyle"/>/css<xsl:value-of select="//MpStyle/Article/Caption"/>/<xsl:value-of select="//NowTime"/>.css</xsl:attribute>
	</link>

    <link rel="stylesheet" href="/knowledge/style/pedia.css" type="text/css" />
    <!--    
    <script src="/js/pedia.js" type="text/javascript">/**/</script>
    -->    
    <script type="text/javascript" src="/subject/aspSrc.asp" >/**/</script>
    <script language="javascript">
        <xsl:attribute name="type">text/javascript</xsl:attribute>
		<xsl:attribute name="src">/xslgip/<xsl:value-of select="myStyle"/>/AC_RunActiveContent.js</xsl:attribute>
        <xsl:text>/**/</xsl:text>
	</script>    
	<script language="javascript">
        <xsl:attribute name="type">text/javascript</xsl:attribute>
		<xsl:attribute name="src">/xslgip/<xsl:value-of select="myStyle"/>/ftiens4.js</xsl:attribute>
        <xsl:text>/**/</xsl:text>
	</script>
  <!-- Added By Leo  2011-6-16  加入jQuery SimplyScroll  Start -->  
  <script language="javascript" type="text/javascript" src="/js/simplyscroll/jquery.simplyscroll-1.0.4.js">/**/</script>
  <link rel="stylesheet" type="text/css" href="/js/simplyscroll/jquery.simplyscroll-1.0.4.css"></link>
  <script type="text/javascript">
      (function($) {
      $(function() { //on DOM ready
      $("#scroller").simplyScroll({
      className: 'vert vert_<xsl:value-of select="//MpStyle/Article/Caption"/>_<xsl:value-of select="//NowTime"/>',
    horizontal: false,
    frameRate: 60,
    speed: 5
    });
    <!-- Added By Leo  2011-07-07 加入jQuery easing  Start  -->

      //get the default top value
      var top_val = $('#path_menu li a').css('top');

      //animate the selected menu item
      $('#path_menu li.selected').children('a').stop().animate({ top: 5 }, { easing: 'easeOutQuad', duration: 200 });

      $('#path_menu li').hover(
      function () {
      //animate the menu item with 0 top value
      $(this).children('a').stop().animate({ top: 5 }, { easing: 'easeOutQuad', duration: 200 });
      },
      function () {
      //set the position to default
      $(this).children('a').stop().animate({ top: top_val }, { easing: 'easeOutQuad', duration: 200 });

      //always keep the previously selected item in fixed position
      $('#path_menu li.selected').children('a').stop().animate({ top: 5 }, { easing: 'easeOutQuad', duration: 200 });
      });
      
      /*
      */
      <!-- Added By Leo  2011-07-07 加入jQuery easing   End  -->
      <!-- Added By Leo  2011-07-07 加入Menu Color Change   Start  -->
      var activeCSSName = 'leftActiveItem_<xsl:value-of select="//MpStyle/Article/Caption"/>_<xsl:value-of select="//NowTime"/>';
      var ctNode = $.query.get('ctnode');
      var CategoryId = $.query.get('categoryid');
      if (ctNode.length != 0)
      {
      // 擁有ctNode的QueryString
      // var develop = $("#農業百年發展史").attr("href");
      // var developNum = develop.indexOf('ctNode=');
      // var developNode = develop.substring(developNum+7,developNum+11);

      var news = $("#最新消息").attr("href");
      var newsNum = news.indexOf('ctNode=');
      var newsNode = news.substring(newsNum+7,newsNum+11);

      var excellent = $("#優質農業人").attr("href");
      var excellentNum = excellent.indexOf('ctNode=');
      var excellentNode = excellent.substring(excellentNum+7,excellentNum+11);

      var life = $("#農業與生活").attr("href");
      var lifeNum = life.indexOf('ctNode=');
      var lifeNode = life.substring(lifeNum+7,lifeNum+11);

      var column = $("#產銷專欄").attr("href");
      var columnNum = column.indexOf('ctNode=');
      var columnNode = column.substring(columnNum+7,columnNum+11);

      var resource = $("#資源推薦").attr("href");
      var resourceNum = resource.indexOf('ctNode=');
      var resourceNode = resource.substring(resourceNum+7,resourceNum+11);

      var vedio = $("#影音專區").attr("href");
      var vedioNum = vedio.indexOf('ctNode=');
      var vedioNode = vedio.substring(vedioNum+7,vedioNum+11);

      var website = $("#相關網站").attr("href");
      var websiteNum = website.indexOf('ctNode=');
      var websiteNode = website.substring(websiteNum+7,websiteNum+11);

      //if (developNode == ctNode)
      //{
      //$("#農業百年發展史").addClass(activeCSSName);
      //}
      //else
      if (newsNode == ctNode)
      {
      $("#最新消息").addClass(activeCSSName);
      }
      else if (excellentNode == ctNode)
      {
      $("#優質農業人").addClass(activeCSSName);
      }
      else if (lifeNode == ctNode)
      {
      $("#農業與生活").addClass(activeCSSName);
      }
      else if (columnNode == ctNode)
      {
      $("#產銷專欄").addClass(activeCSSName);
      }
      else if (resourceNode == ctNode)
      {
      $("#資源推薦").addClass(activeCSSName);
      }
      else if (vedioNode == ctNode)
      {
      $("#影音專區").addClass(activeCSSName);
      }
      else if (websiteNode == ctNode)
      {
      $("#相關網站").addClass(activeCSSName);
      }
      }
      else if (CategoryId.length != 0)
      {
      switch (CategoryId)
      {
      case "a":
      $("#農").addClass(activeCSSName);
      break;
      case "b":
      $("#林").addClass(activeCSSName);
      break;
      case "c":
      $("#漁").addClass(activeCSSName);
      break;
      case "d":
      $("#牧").addClass(activeCSSName);
      break;
      case "e":
      $("#其它").addClass(activeCSSName);
      break;
      default:
      $("#全部").addClass(activeCSSName);
      break;
      }
      }
      else
      {
      var search_recommand = location.href.search("recommand");
      var search_jigsaw2010 = location.href.search("jigsaw2010");
      var search_century = location.href.search("Century");

      if (search_recommand > -1)
      {
      $("#好文推薦").addClass(activeCSSName);
      }
      if (search_jigsaw2010 > -1)
      {
      $("#農漁生產地圖").addClass(activeCSSName);
      }
      if (search_century > -1)
      {
      $("#百年農業發展史").addClass(activeCSSName);
      }
      }
      <!-- Added By Leo  2011-07-07 加入Menu Color Change    End   -->
      });
    })(jQuery);
  </script>
  <!-- Added By Leo  2011-6-16  加入jQuery SimplyScroll  End  -->
  
  <!-- Added By Leo  2011-07-07 加入jQuery easing  Start  -->
  <script language="javascript" type="text/javascript" src="/js/jqueryeasing/js/jquery.easing.1.3.js">/**/</script>
  <link rel="stylesheet" type="text/css" href="/js/jqueryeasing/css/jQueryEasing.css"></link>
  <!-- Added By Leo  2011-07-07 加入jQuery easing   End  -->
  <!-- Added By Leo  2011-07-07 加入Menu Color Change   Start  -->
  <script language="javascript" type="text/javascript" src="/js/jquery.query-2.1.7.js">/**/</script>
  <link rel="stylesheet" type="text/css" href="/css/menu_Selected.css" />
  <!-- Added By Leo  2011-07-07 加入Menu Color Change    End   -->
  
	<script type="text/javascript">var fornothing = 00000;</script>
	<script type="text/javascript">
		var gaJsHost=(("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
		document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
	</script>
	<script type="text/javascript">
		try{
		var pageTracker=_gat._getTracker("UA-9195501-1");
		pageTracker._trackPageview();
		}catch(err)
		{}
	</script>

</xsl:template>
<!-- Header／結束 -->

<!-- 會員字型設定共用變數-->
	<xsl:variable name="fontsize">
		<xsl:choose>
			<xsl:when test="//login/memFontSize='12'">12</xsl:when>
			<xsl:when test="//login/memFontSize='14'">14</xsl:when>
			<xsl:when test="//login/memFontSize='16'">16</xsl:when>
			<xsl:when test="//login/memFontSize='18'">18</xsl:when>
			<xsl:otherwise>12</xsl:otherwise>
		</xsl:choose>
	</xsl:variable>
<!-- 會員字型設定結束-->

<!-- TopLink／開始 -->
<xsl:template match="hpMain" mode="toplink">
	<ul class="nav">
		<xsl:apply-templates select="." mode="n"/>
		<xsl:for-each select="TopLink/Article">
			<li>
				<a>
					<xsl:attribute name="href"><xsl:value-of select="xURL" /></xsl:attribute>
					<xsl:if test="@newWindow='Y'"><xsl:attribute name="target">_blank</xsl:attribute></xsl:if>
					<xsl:attribute name="title"><xsl:value-of select="Caption"/></xsl:attribute>
					<xsl:value-of select="Caption"/>
				</a>
			</li>
		</xsl:for-each>
	</ul>
</xsl:template>
<!-- TopLink／結束 -->

<!-- 分網連結／開始 -->
<xsl:template match="hpMain" mode="weblink">
	<div class="subnav">
  <ul>
		<xsl:for-each select="WebLink/Article">
			<li>
				<a>
					<xsl:attribute name="href"><xsl:value-of select="xURL" /></xsl:attribute>
					<xsl:if test="@newWindow='Y'"><xsl:attribute name="target">_blank</xsl:attribute></xsl:if>
					<xsl:attribute name="title"><xsl:value-of select="Caption"/></xsl:attribute>
					<xsl:value-of select="Caption"/>
				</a>
			</li>
		</xsl:for-each>
	</ul>
	</div>
</xsl:template>
<!-- 分網連結／結束 -->

<!-- RunActive／開始 -->
<xsl:template match="hpMain" mode="RunActive">
	<div>
		<xsl:attribute name="class">header</xsl:attribute>
		<xsl:variable name="url">/xslgip/<xsl:value-of select="myStyle"/>/images<xsl:value-of select="//MpStyle/Article/Caption"/></xsl:variable>
	  <h1><a href="#"><xsl:value-of select="$title"/></a></h1>
		<script language="javascript">
			if (AC_FL_RunContent == 0) {
				alert("這個頁面必須具備 AC_RunActiveContent.js。");
			} else {
				AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,0,0','width','980','height','141','id','banner','align','middle','name','banner','src','<xsl:value-of select="$url"/>/banner<xsl:value-of select="//NowTime"/>','quality','high','wmode','transparent','allowscriptaccess','sameDomain','allowfullscreen','false','pluginspage','http://www.macromedia.com/go/getflashplayer','movie','<xsl:value-of select="$url"/>/banner<xsl:value-of select="//NowTime"/>' ); //end AC code
			}
		</script>
		<noscript>
			<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,0,0" name="banner" width="980" height="141" align="middle" id="banner">
				<param name="allowScriptAccess" value="sameDomain" />
				<param name="allowFullScreen" value="false" />
				<param name="quality" value="high" />
				<param name="wmode" value="transparent" />
				<param name="movie">
					<xsl:attribute name="value">/xslgip/<xsl:value-of select="myStyle"/>/images<xsl:value-of select="//MpStyle/Article/Caption"/>/banner<xsl:value-of select="//NowTime"/>.swf</xsl:attribute>
				</param>
				<embed quality="high" wmode="transparent" width="980" height="141" name="banner" align="middle" allowScriptAccess="sameDomain" allowFullScreen="false" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer">
					<xsl:attribute name="src">/xslgip/<xsl:value-of select="myStyle"/>/images<xsl:value-of select="//MpStyle/Article/Caption"/>/banner<xsl:value-of select="//NowTime"/>.swf</xsl:attribute>
				</embed>
			</object>
		</noscript>
	</div>
		<div id="cursor" style="height: 25px; border: left: 100px; position: absolute; z-index:1000; top: 160px; visibility: visible; width: 10px">
		<img id="cursor-img" style="display: none" />
		</div>
		
		<script language="JavaScript">
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
			<![CDATA[
			if((cursorInnerWidth-55)<=x){
				x = cursorInnerWidth -55;}
			if((cursorInnerHeight -45)<=y ){
				y = cursorInnerHeight -45;}]]>
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
		</script>
</xsl:template>
<!-- RunActive／結束 -->

<!-- 節氣／開始 -->
<xsl:template match="hpMain" mode="SolarTerms">
	<div class="season">
		<p>
      <a href="/DateTransfer.aspx"><xsl:value-of select="//today"/></a>
    </p>
    <p><xsl:value-of select="//BlockG/Article/Caption"/></p>
		<p><span>
			<xsl:choose>
				<xsl:when test="//BlockG/Article/Content=''"></xsl:when>
				<xsl:otherwise><xsl:value-of disable-output-escaping="yes" select="//BlockG/Article/Content"/></xsl:otherwise>
			</xsl:choose>
		</span></p>
		<h2><xsl:value-of select="SolarTerms/Article/Caption"/></h2>
		<p class="statement">
			<a>
				<xsl:attribute name="href">/ct.asp?xItem=<xsl:value-of select="//SolarTerms/Article/@iCuItem" />&amp;ctNode=<xsl:value-of select="//SolarTerms/@xNode" /></xsl:attribute>
				<xsl:attribute name="title"><xsl:value-of disable-output-escaping="yes" select="substring(//SolarTerms/Article/Content,1,30)" /></xsl:attribute>
				<xsl:value-of disable-output-escaping="yes" select="substring(//SolarTerms/Article/Content,1,20)" />...
			</a>
		</p>
	</div>
</xsl:template>
<!-- 節氣／結束 -->

<!-- 選單／開始 -->
<xsl:template match="hpMain" mode="TreeScriptCode">
	<xsl:value-of disable-output-escaping="yes" select="TreeScriptCode" />
</xsl:template>

<xsl:template match="hpMain" mode="Menu">
  <div class="menu">
		<UL>
			<xsl:for-each select="MenuBar1/MenuCat">
				<li>
					<a>
						<xsl:choose>
							<xsl:when test="contains(redirectURL, 'http')">
								<xsl:attribute name="href"><xsl:value-of select="redirectURL" /></xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="href">/<xsl:value-of select="redirectURL" /></xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
						<xsl:if test="@newWindow='Y'"><xsl:attribute name="target">_blank</xsl:attribute></xsl:if>
						<xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute>
            <xsl:attribute name="id"><xsl:value-of select="Caption" /></xsl:attribute>
						<xsl:value-of select="Caption" />
					</a>
					<xsl:if test="MenuItem">
						<ul>
							<xsl:for-each select="MenuItem">
								<xsl:apply-templates select="." />
							</xsl:for-each>
						</ul>
					</xsl:if>
				</li>
			</xsl:for-each>
		</UL>
	</div>
</xsl:template>

<xsl:template match="MenuItem">
	<li>
		<a>
			<xsl:choose>
				<xsl:when test="contains(redirectURL, 'http')">
					<xsl:attribute name="href"><xsl:value-of select="redirectURL" /></xsl:attribute>
				</xsl:when>
				<xsl:otherwise>
					<xsl:attribute name="href">/<xsl:value-of select="redirectURL" /></xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>
			<xsl:if test="@newWindow='Y'"><xsl:attribute name="target">_blank</xsl:attribute></xsl:if>
			<xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute>
			<xsl:value-of select="Caption" />
		</a>
	</li>
</xsl:template>
<!-- 選單／結束 -->

<!-- RSS Start-->
<xsl:template match="hpMain" mode="xsRSS">
	<div class="rss">
		<a href="lp.asp?ctNode=1574&amp;CtUnit=305&amp;BaseDSD=7&amp;mp=1">RSS訂閱服務</a>
	</div>
</xsl:template>
<!-- RSS End-->

<!-- AD Start-->
<xsl:template match="hpMain" mode="xsAD">
    <div class="ad">
        <ul id="scroller">
            <xsl:for-each select="//AD/Article">
                <li>
                    <a>
                        <xsl:attribute name="href">
                            <xsl:value-of select="xURL" />
                        </xsl:attribute>
                        <xsl:if test="@newWindow='Y'">
                            <xsl:attribute name="target">_blank</xsl:attribute>
                        </xsl:if>
                        <img>
                            <xsl:attribute name="src">
                                /<xsl:value-of select="xImgFile" />
                            </xsl:attribute>
                            <xsl:attribute name="alt">
                                <xsl:value-of select="Caption" />
                            </xsl:attribute>
                        </img>
                    </a>
                </li>
            </xsl:for-each>
        </ul>
    </div>
</xsl:template>
<!-- AD End-->

<!--路徑連結 Start-->
<xsl:template match="hpMain" mode="xsPath">
  <div class="path" style="float:left; padding-top: 10px">目前位置：</div>
  <div style="float:left;width:80%">
  <ul id="path_menu">
    <li>
     <a title="首頁">
        <xsl:attribute name="href">
          /mp.asp?mp=<xsl:value-of select="//mp" />
        </xsl:attribute>
        首頁
      </a>
    </li>
    <xsl:for-each select="//xPath/xPathNode">
      <li style="top: 10px;">></li>
      <li><a>
        <xsl:attribute name="href">
          /np.asp?ctNode=<xsl:value-of select="@xNode" />&amp;mp=<xsl:value-of select="//mp" />
        </xsl:attribute>
        <!--<xsl:attribute name="title">
          <xsl:value-of select="@Title" />
        </xsl:attribute>-->
        <xsl:value-of select="@Title" />
      </a>
      </li>
    </xsl:for-each>
    <xsl:for-each select="//CatList/mediapath">
      <xsl:if test="mcode!=''">
        <xsl:if test="mvalue!=''">
        <li style="top: 10px;">></li>
        <xsl:choose>
          <xsl:when test="murl!=''">
            <li><a>
              <xsl:attribute name="href">
                <xsl:value-of select="murl" />&amp;mp=<xsl:value-of select="//mp" />
              </xsl:attribute>
              <xsl:attribute name="title">
                <xsl:value-of select="mvalue" />
              </xsl:attribute>
              <xsl:value-of select="mvalue" />
            </a>
            </li>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="mvalue" />
          </xsl:otherwise>
        </xsl:choose>
        </xsl:if>
      </xsl:if>
    </xsl:for-each>
    <!--路徑套用主分類 START-->
    <xsl:for-each select="//CatList/CatItem">
      <xsl:if test="contains(concat(//qURL, ''), concat(xqCondition, ''))">
        <li style="top: 10px;">></li>
      </xsl:if>
      <xsl:if test="contains(concat(//qURL, ''), concat(xqCondition, ''))">
        <li><a>
          <xsl:attribute name="href">
            lp.asp?<xsl:value-of select="../../xqURL" /><xsl:value-of select="xqCondition" />
          </xsl:attribute>
          <xsl:attribute name="title">
            <xsl:value-of select="CatName" />
          </xsl:attribute>
          <xsl:value-of select="CatName" />
        </a>
        </li>
      </xsl:if>
    </xsl:for-each>
    <!--路徑套用主分類 END-->
  </ul>
	</div>
</xsl:template>
<!--路徑連結 End-->

<!--頁面功能 Start-->
<xsl:template match="hpMain" mode="xsFunction">
	<div class="Function2">
    <ul>	
			<li>
				<a class="Print">
					<xsl:attribute name="href">#</xsl:attribute>
					<xsl:attribute name="onclick">window.open('/fp.asp?xItem=<xsl:value-of select="//@iCuItem" />&amp;ctNode=<xsl:value-of select="//xPathNode[position()=last()]/@xNode" />&amp;mp=<xsl:value-of select="//mp" />')</xsl:attribute>
					<xsl:attribute name="title">友善列印</xsl:attribute>友善列印
				</a>
			</li>
			<!--<li>
				<a href="javascript:history.go(-1);" class="Back" title="回上一頁">回上一頁</a>
				<noscript>本網頁使用SCRIPT編碼方式執行回上一頁的動作，如果您的瀏覽器不支援SCRIPT，請直接使用「Alt+方向鍵向左按鍵」來返回上一頁</noscript>
			</li>-->
        <li>
        <a id="a_list" href="#" class="List" title="回列表頁">回列表頁</a>
      </li>
      <li>
        <xsl:if test="BackURL!=''">
          <a class="Back" title="上一篇">
            <xsl:attribute name="href">
              <xsl:value-of select="BackURL"/>
            </xsl:attribute>上一篇
          </a>
        </xsl:if>
      </li>
      <li>
        <xsl:if test="NextURL!=''">
          <a class="Next" title="下一篇">
            <xsl:attribute name="href">
              <xsl:value-of select="NextURL"/>
            </xsl:attribute>下一篇
          </a>
        </xsl:if>
      </li>
		</ul>
	</div>
  <a>
    <xsl:value-of select="ArticleNavSQLScript"/>
  </a>
			<DIV align="right">
        
                    調整字級：
                   
                    <A onclick="changeFontSize('article','idx1');">
                      <IMG id="idx1" alt="" src="subject/images/fontsize_1_off.gif" align="absMiddle" border="0" name="idx1" style="padding-left:2px;padding-right:2px;padding-bottom:2px;" />
                    </A>
                    <A onclick="changeFontSize('article','idx2');">
                      <IMG id="idx2" alt="" src="subject/images/fontsize_2_off.gif" align="absMiddle" border="0" name="idx2" style="padding-left:2px;padding-right:2px;padding-bottom:2px;" />
                    </A>
                    <A onclick="changeFontSize('article','idx3');">
                      <IMG id="idx3" alt="" src="subject/images/fontsize_3_off.gif" align="absMiddle" border="0" name="idx3" style="padding-left:2px;padding-right:2px;padding-bottom:2px;" />
                    </A>
                    <A onclick="changeFontSize('article','idx4');">
                      <IMG id="idx4" alt="" src="subject/images/fontsize_4_off.gif" align="absMiddle" border="0" name="idx4" style="padding-left:2px;padding-right:2px;padding-bottom:2px;" />
                    </A>
					
                  </DIV>
				 
</xsl:template>
<!--頁面功能 End-->

<!--電子報 Start-->
<xsl:template match="hpMain" mode="xsePaper">
	<div class="epaper">
		<h2>訂閱電子報</h2>
		<div class="body">
			<form method="post" action="/epaper_act.asp?CtRootID=21">
				<input name="textfield2" type="text" size="20" class="txt" value="請輸入電子郵件帳號" onfocus="JSCRIPT: this.value='';"/>
				<input name="submit1" type="image" alt="訂閱" class="btn">
					<xsl:attribute name="src">/xslgip/<xsl:value-of select="myStyle"/>/images/epaperBtn1.gif</xsl:attribute>
				</input>
				<input name="submit2" type="image" alt="取消訂閱" class="btn">
					<xsl:attribute name="src">/xslgip/<xsl:value-of select="myStyle"/>/images/epaperBtn2.gif</xsl:attribute>
				</input>
      <a href="/lp.asp?CtNode=2873&amp;CtUnit=1356&amp;BaseDSD=7&amp;mp=1">
      <img alt="歷史電子報" border="0">
					<xsl:attribute name="src">/xslgip/<xsl:value-of select="myStyle"/>/images/epaperBtn3.gif</xsl:attribute>
      </img>
      </a>
			</form>
		</div>
	</div>
</xsl:template>
<!--電子報 End-->

<!--搜尋服務 Start-->
<xsl:template match="hpMain" mode="xsSearch">
	<div class="search">
		<h2>搜尋服務</h2>
		<form name="SearchForm" method="post" taget="_self" action="">
		<!-- Modified By Leo 2011-09-08	更換搜尋引擎	Start	-->
		<!--<table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td colspan="2">
            <input name="Keyword" type="text" accesskey="s" class="txt" onkeydown="javascript:checkKeyDown(event)" />
          </td>
        </tr>
        <tr>
          <td>
            <input name="FromSiteUnit" value="1" type="checkbox" class="ckb" checked="checked" />站內單元
          </td>
          <td>
            <input name="FromKnowledgeTank" value="1" type="checkbox" class="ckb" checked="checked" />知識庫
          </td>
        </tr>
        <tr>
          <td>
            <input name="FromKnowledgeHome" value="1" type="checkbox" class="ckb" checked="checked" />知識家
          </td>
          <td>
            <input name="FromTopic" value="1" type="checkbox" class="ckb"  checked="checked" />主題館
          </td>
        </tr>
        <tr>
          <td>
            <input name="FromPedia" value="1" type="checkbox" class="ckb"  checked="checked" />農業小百科
          </td>
          <td>
            <input name="FromVideo" value="1" type="checkbox" class="ckb"  checked="checked" />影音專區
          </td>
        </tr>
        <tr>
          <td colspan="2">
            <input name="FromTechCD" value="1" type="checkbox" class="ckb"  checked="checked" />技術光碟
          </td>
        </tr>
        </table>-->
      <table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td colspan="2">
            <input name="Keyword" type="text" accesskey="s" class="txt" onkeydown="javascript:checkKeyDown(event)" />
          </td>
        </tr>
        <tr>
          <td>
            <input name="FromSiteUnit" value="1" type="checkbox" class="ckb" checked="checked" />站內單元
          </td>
          <td>
            <input name="FromSiteUnit" value="7" type="checkbox" class="ckb" checked="checked" />知識庫
          </td>
        </tr>
        <tr>
          <td>
            <input name="FromSiteUnit" value="3" type="checkbox" class="ckb" checked="checked" />知識家
          </td>
          <td>
            <input name="FromSiteUnit" value="5" type="checkbox" class="ckb"  checked="checked" />主題館
          </td>
        </tr>
        <tr>
          <td>
            <input name="FromSiteUnit" value="4" type="checkbox" class="ckb"  checked="checked" />農業小百科
          </td>
          <td>
            <input name="FromSiteUnit" value="2" type="checkbox" class="ckb"  checked="checked" />影音專區
          </td>
        </tr>
        <tr>
          <td>
            <input name="FromSiteUnit" value="6" type="checkbox" class="ckb"  checked="checked" />技術光碟
          </td>
          <td>
            <input name="FromSiteUnit" value="11" type="checkbox" class="ckb"  checked="checked" />典藏豐年
          </td>
        </tr>
        </table>
		<!-- Modified By Leo 2011-09-08	更換搜尋引擎	 End 	-->
	  	<input id="btnSubmit" name="search" type="image" alt="search" class="btn" onClick="javascript:checkSearchForm(0)">
       	<xsl:attribute name="src">/xslgip/<xsl:value-of select="myStyle"/>/images/searchBtn.gif</xsl:attribute>
      </input>
	  	<input name="search" type="image" alt="search" class="btn" onClick="javascript:checkSearchForm(1)">
       	<xsl:attribute name="src">/xslgip/<xsl:value-of select="myStyle"/>/images/SearchBtn2.gif</xsl:attribute>
      </input>
		</form>
		<script language="javascript">
		<![CDATA[

			function checkSearchForm(value)
			{
				if( value == 0 ) {
					if( document.SearchForm.Keyword.value.replace(/^[\s　]+/g, "").replace(/[\s　]+$/g, "") == "" ) {
						alert('請輸入查詢值');
						document.SearchForm.Keyword.value="";
						event.returnValue = false;
					}
					else {
						document.SearchForm.Keyword.value = document.SearchForm.Keyword.value.replace(/^[\s　]+/g, "").replace(/[\s　]+$/g, "");
						//document.SearchForm.action = "/kp.asp?xdURL=Search/SearchResultList.asp";
            document.SearchForm.action = "/Search/SearchResult.aspx";
						document.SearchForm.submit();
					}
				}
				else {
					document.SearchForm.Keyword.value = document.SearchForm.Keyword.value.replace(/^[\s　]+/g, "").replace(/[\s　]+$/g, "");
					document.SearchForm.action = "/kp.asp?xdURL=Search/AdvancedSearch.asp";
					document.SearchForm.submit();
				}
			}

			function checkKeyDown(event)
			{
				if(event.which || event.keyCode)
				{
					if ((event.which == 13) || (event.keyCode == 13))
					{
						document.getElementById('btnSubmit').click();
						return false;
					}
				}
				else
				{
					return true;
				}
			}
		]]>
		</script>
	</div>
</xsl:template>
<!--搜尋服務 End-->

<!--  relatedDocument -->
<xsl:template match="hpMain" mode="relatedDocument">
	<div class="rbox">
		<h2>關聯文章</h2>
		<div class="body">
			<xsl:if test="//relatedDocument">
				<xsl:choose>
					<xsl:when test="//relatedDocument/isConnect='Y' and //relatedDocument/haveResult='Y'">
						<ul>
						<xsl:for-each select="//relatedDocument/Article">
							<xsl:choose>
								<xsl:when test="stitle=''">
								</xsl:when>
								<xsl:otherwise>
									<li>
										<a>
											<xsl:attribute name="href"><xsl:value-of select="url" /></xsl:attribute>
											<xsl:attribute name="title"><xsl:value-of select="stitle" /></xsl:attribute>
											<xsl:value-of select="stitle" />
										</a>
									</li>
								</xsl:otherwise>
							</xsl:choose>
						</xsl:for-each>
						</ul>
					</xsl:when>
					<xsl:otherwise>
						<ul><li>無相關聯文章</li></ul>
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>
		</div>
  <div class="Foot"></div>
	</div>
</xsl:template>

<!--頁尾資訊 Start-->
<xsl:template match="hpMain" mode="xsFooter">
	<xsl:choose>
		<xsl:when test="//Footer/Article/xImgFile">
			<img>
				<xsl:attribute name="src">/<xsl:value-of select="//Footer/Article/xImgFile" /></xsl:attribute>
				<xsl:attribute name="alt"><xsl:value-of select="//Footer/Article/Caption" /></xsl:attribute>
			</img>
		</xsl:when>
		<xsl:otherwise>
			<img>
				<xsl:attribute name="src">/xslgip/<xsl:value-of select="//myStyle" />/images/webindoor.gif</xsl:attribute>
				<xsl:attribute name="alt"><xsl:value-of select="//Footer/Article/Caption" /></xsl:attribute>
			</img>
		</xsl:otherwise>
	</xsl:choose>
	<p><xsl:value-of disable-output-escaping="yes"  select="//Footer/Article/ContentAll" /></p>
	<div class="count">累積瀏覽人次：<xsl:value-of select="//Counter" /><br/>
	
	</div>
	<script type="text/javascript" src="/ViewCounter.aspx">/**/</script>
</xsl:template>

<xsl:template match="hpMain" mode="xsCopyright">
	<div class="copyright">
		<xsl:apply-templates select="." mode="b"/>
		<xsl:for-each select="//Copyright/Article">
			<a>
				<xsl:attribute name="href"><xsl:value-of select="xURL" /></xsl:attribute>
				<xsl:if test="@newWindow='Y'"><xsl:attribute name="target">_blank</xsl:attribute></xsl:if>
				<xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute>
				<xsl:value-of select="Caption" />
			</a>
			<xsl:if test="position()!=last()">｜</xsl:if>
		</xsl:for-each>
	</div>
</xsl:template>
<!--頁尾資訊 End-->


<!-- Accesskey B／開始 -->
<xsl:template match="hpMain" mode="b">
	<a title="網頁下方群組連結" class="Accesskey" accesskey="B">
		<xsl:apply-templates select="." mode="link" />
	</a>
</xsl:template>
<!-- Accesskey B／結束 -->

<!-- Accesskey C／開始 -->
<xsl:template match="hpMain" mode="c">
	<a title="網頁內容資料" class="Accesskey" accesskey="C">
		<xsl:apply-templates select="." mode="link" />
	</a>
</xsl:template>
<!-- Accesskey C／結束 -->


<!-- Accesskey L／開始 -->
<xsl:template match="hpMain" mode="l">
	<a title="左方主選單區" class="Accesskey" accesskey="L">
		<xsl:apply-templates select="." mode="link" />
	</a>
</xsl:template>
<!-- Accesskey L／結束 -->

<!-- Accesskey N／開始 -->
<xsl:template match="hpMain" mode="n">
	<a title="網頁上方導覽連結區" class="Accesskey" accesskey="N">
		<xsl:apply-templates select="." mode="link" />
	</a>
</xsl:template>
<!-- Accesskey N／結束 -->

<!-- Accesskey R／開始 -->
<xsl:template match="hpMain" mode="r">
	<a title="網頁右方欄" class="Accesskey" accesskey="R">
		<xsl:apply-templates select="." mode="link" />
	</a>
</xsl:template>
<!-- Accesskey R／結束 -->

<!--  AccesskeyLink／開始 -->
<xsl:template match="hpMain" mode="link">
<xsl:attribute name="href">
	/ct.asp?xItem=83771&amp;ctNode=1592&amp;mp=<xsl:value-of select="//mp" />
</xsl:attribute>:::
</xsl:template>
<!--  AccesskeyLink／結束 -->


<!-- 聯絡我們表單／開始 -->
<xsl:template match="hpMain" mode="forward">
	<xsl:copy-of select="//SCRIPT"/>
	<form class="Contactus" name="XForm" onSubmit="return validateForm(document.XForm);"><xsl:attribute name="action"><xsl:value-of select="//Action" /></xsl:attribute>
		<table>
			<caption><xsl:value-of select="MainArticle /Caption" /></caption>
			<input type="hidden" name="CtUnit">
				<xsl:attribute name="value"><xsl:value-of select="CtUnit" /></xsl:attribute>
			</input>
			<input type="hidden" name="xItem">
				<xsl:attribute name="value"><xsl:value-of select="xItem" /></xsl:attribute>
			</input>
			<input type="hidden" name="ctNode">
				<xsl:attribute name="value"><xsl:value-of select="ctNode" /></xsl:attribute>
			</input>
			<input type="hidden" name="mp">
				<xsl:attribute name="value"><xsl:value-of select="mp" /></xsl:attribute>
			</input>
			<input type="hidden" name="TopInfo">
				<xsl:attribute name="value"><xsl:value-of select="TopInfo" /></xsl:attribute>
			</input>
			<xsl:apply-templates select="//FieldList" />
		</table>
		<xsl:copy-of select="//pHTML"/>
	</form>
</xsl:template>
<!-- 聯絡我們表單／結束 -->

<xsl:template match="FieldList">
	<xsl:for-each select="Field">
		<tr>
			<th scope="row"><xsl:value-of select="Title" /></th>
			<td>
				<xsl:if test="Type='text'">
					<input type="text" class="Text">
						<xsl:attribute name="name"><xsl:value-of select="Name" /></xsl:attribute>
						<xsl:attribute name="size"><xsl:value-of select="Size" /></xsl:attribute>
						<xsl:attribute name="value"><xsl:value-of select="Value" /></xsl:attribute>
						<xsl:attribute name="onfocus">JSCRIPT: if(this.value=='<xsl:value-of select="Value" />'){this.value='';}</xsl:attribute>
					</input>
				</xsl:if>
				<xsl:if test="Type='label'">
					<label>
						<xsl:value-of select="Value" />
					</label>
				</xsl:if>
				<xsl:if test="Type='textarea'">
					<textarea wrap="VIRTUAL">
						<xsl:attribute name="name"><xsl:value-of select="Name" /></xsl:attribute>
						<xsl:attribute name="cols"><xsl:value-of select="Cols" /></xsl:attribute>
						<xsl:attribute name="rows"><xsl:value-of select="Rows" /></xsl:attribute>
						<xsl:attribute name="onfocus">JSCRIPT: if(this.value=='<xsl:value-of select="Value" />'){this.value='';}</xsl:attribute>
						<xsl:value-of select="Value" />
					</textarea>
				</xsl:if>
			</td>
		</tr>
	</xsl:for-each>
</xsl:template>
</xsl:stylesheet>
