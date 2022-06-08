<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:hyweb="urn:gip-hyweb-com" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
<xsl:include href="pediamember.xsl"/>
<xsl:include href="../hwXslFunc.xsl"/>
<xsl:variable name="title">農業知識入口網</xsl:variable>	


<!-- Header／開始 -->
<xsl:template match="hpMain" mode="Header">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title><xsl:value-of select="$title"/>－農業知識小百科－</title>
	<link rel="stylesheet" type="text/css">
		<xsl:attribute name="href">../xslgip/<xsl:value-of select="myStyle"/>/css<xsl:value-of select="//MpStyle/Article/Caption"/>/common.css</xsl:attribute>
	</link>
	<link rel="stylesheet" type="text/css">
		<xsl:attribute name="href">../xslgip/<xsl:value-of select="myStyle"/>/css<xsl:value-of select="//MpStyle/Article/Caption"/>/memberMOF.css</xsl:attribute>
	</link>
	<script language="javascript">AC_FL_RunContent = 0;</script>
	<script src="../xslgip/style1/AC_RunActiveContent.js" language="javascript">AC_FL_RunContent = 0;</script>
	<script src="../xslgip/style1/ftiens4.js" language="javascript">AC_FL_RunContent = 0;</script>
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
	<ul class="subnav">
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
</xsl:template>
<!-- 分網連結／結束 -->

<!-- RunActive／開始 -->
<xsl:template match="hpMain" mode="RunActive">
	<div>
		<xsl:attribute name="class">header<xsl:value-of select="//NowTime"/></xsl:attribute>
		<xsl:variable name="url">../xslgip/<xsl:value-of select="myStyle"/>/images<xsl:value-of select="//MpStyle/Article/Caption"/></xsl:variable>
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
				<xsl:attribute name="value">../xslgip/<xsl:value-of select="myStyle"/>/images<xsl:value-of select="//MpStyle/Article/Caption"/>/banner<xsl:value-of select="//NowTime"/>.swf</xsl:attribute>
			</param>
			<embed quality="high" wmode="transparent" width="980" height="141" name="banner" align="middle" allowScriptAccess="sameDomain" allowFullScreen="false" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer">
				<xsl:attribute name="src">../xslgip/<xsl:value-of select="myStyle"/>/images<xsl:value-of select="//MpStyle/Article/Caption"/>/banner<xsl:value-of select="//NowTime"/>.swf</xsl:attribute>
			</embed>
		  </object>
		</noscript>
	
	</div>
</xsl:template>
<!-- RunActive／結束 -->

<!-- 節氣／開始 -->
<xsl:template match="hpMain" mode="SolarTerms">
	<div class="season">
		<h2><xsl:value-of select="SolarTerms/Article/Caption"/></h2>
		<p>
			<span>
				<xsl:choose>
					<xsl:when test="//BlockG/Article/Content=''"><br/><br/></xsl:when>
					<xsl:otherwise><xsl:value-of disable-output-escaping="yes" select="//BlockG/Article/Content"/></xsl:otherwise>
				</xsl:choose>
			</span>
		</p>
		<p>
			<xsl:value-of select="//today"/><!--xsl:value-of select="hyweb:edayOfWeek(etoday)"/--><br/><xsl:value-of select="//BlockG/Article/Caption"/>
		</p>
		<p>
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
	<UL class="menu">
		<xsl:for-each select="MenuBar1/MenuCat">
			<li>
				<a>
					<xsl:attribute name="href">/<xsl:value-of select="redirectURL" /></xsl:attribute>
					<xsl:if test="@newWindow='Y'"><xsl:attribute name="target">_blank</xsl:attribute></xsl:if>
					<xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute>
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
</xsl:template>
<xsl:template match="MenuItem">
	<li>
		<a>
			<xsl:attribute name="href">/<xsl:value-of select="redirectURL" /></xsl:attribute>
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
		<a href="/rss.asp">RSS訂閱服務</a>
	</div>
</xsl:template>
<!-- RSS End-->


<!-- AD Start-->
<xsl:template match="hpMain" mode="xsAD">
<div class="ad">
	<ul>
		<xsl:for-each select="//AD/Article">
			<li>
				<a>
					<xsl:attribute name="href"><xsl:value-of select="xURL" /></xsl:attribute>
					<xsl:if test="@newWindow='Y'"><xsl:attribute name="target">_blank</xsl:attribute></xsl:if>
					<img>
						<xsl:attribute name="src">../<xsl:value-of select="xImgFile" /></xsl:attribute>
						<xsl:attribute name="alt"><xsl:value-of select="Caption" /></xsl:attribute>
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
	<div class="path">目前位置：
		<a title="首頁">
			<xsl:attribute name="href">mp.asp?mp=<xsl:value-of select="//mp" /></xsl:attribute>
			首頁
		</a>
		<xsl:for-each select="//xPath/xPathNode">
			>
			<a>
				<xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="@xNode" />&amp;mp=<xsl:value-of select="//mp" /></xsl:attribute>
				<xsl:attribute name="title"><xsl:value-of select="@Title" /></xsl:attribute>
				<xsl:value-of select="@Title" />
			</a>
		</xsl:for-each>
	</div>
</xsl:template>
<!--路徑連結 End-->


<!--頁面功能 Start-->
<xsl:template match="hpMain" mode="xsFunction">
	<ul class="Function2">
		<!--li><a href="#" class="Track">放入追蹤</a></li-->
		<li>
			<a class="Print">
				<xsl:attribute name="href">fp.asp?xItem=<xsl:value-of select="//@iCuItem" />&amp;ctNode=<xsl:value-of select="//xPathNode[position()=last()]/@xNode" />&amp;mp=<xsl:value-of select="//mp" /></xsl:attribute>
				<xsl:attribute name="title">友善列印</xsl:attribute>
				<xsl:attribute name="target">_blank</xsl:attribute>
				友善列印
			</a>
		</li>
		<!--li>
			<a class="Forward">
				<xsl:attribute name="href">form.asp?xItem=<xsl:value-of select="//@iCuItem" />&amp;ctNode=<xsl:value-of select="//xPathNode[position()=last()]/@xNode" />&amp;mp=<xsl:value-of select="//mp" />&amp;spec=forward</xsl:attribute>
				<xsl:attribute name="title">轉寄好友</xsl:attribute>
				<xsl:attribute name="target">_blank</xsl:attribute>
				轉寄好友
			</a>
		</li-->
		<li>
			<a href="javascript:history.go(-1);" class="Back" title="回上一頁">回上一頁</a>
			<noscript>	本網頁使用SCRIPT編碼方式執行回上一頁的動作，如果您的瀏覽器不支援SCRIPT，請直接使用「Alt+方向鍵向左按鍵」來返回上一頁</noscript>
		</li>
	</ul>
</xsl:template>
<!--頁面功能 End-->

<!--電子報 Start-->
<xsl:template match="hpMain" mode="xsePaper">
	<div class="epaper">
		<h2>訂閱電子報</h2>
		<div class="body">
			<form method="post" action="epaper_act.asp?CtRootID=21">
				<input name="textfield2" type="text" size="20" class="txt" value="請輸入電子郵件帳號" onfocus="JSCRIPT: this.value='';"/>
				<input name="submit1" type="image" alt="訂閱" class="btn">
					<xsl:attribute name="src">../xslgip/<xsl:value-of select="myStyle"/>/images<xsl:value-of select="//MpStyle/Article/Caption"/>/epaperBtn1.gif</xsl:attribute>
				</input>
				<input name="submit2" type="image" alt="取消訂閱" class="btn">
					<xsl:attribute name="src">../xslgip/<xsl:value-of select="myStyle"/>/images<xsl:value-of select="//MpStyle/Article/Caption"/>/epaperBtn2.gif</xsl:attribute>
				</input>
			</form>
		</div>
		<div class="foot"><a href=""></a></div>
	</div>
</xsl:template>
<!--電子報 End-->	

<!--搜尋服務 Start-->
<xsl:template match="hpMain" mode="xsSearch">
	<div class="search">
		<h2>搜尋服務</h2>
		<!--form method="post" taget="hysearch" target="newsearch" action="HTTP://kmsearch.coa.gov.tw/hysearch/cgi/m_query.exe"-->
		<form name="SearchForm" method="post" taget="_self" >
			<label for="search">search</label>
			<input name="Keyword" type="text" accesskey="s" class="txt" />
			<label for="站內單元" class="ckb"><input name="FromSiteUnit" value="1" type="checkbox" class="ckb" checked="checked" />站內單元</label>
			<label for="知識庫" class="ckb"><input name="FromKnowledgeTank" value="1" type="checkbox" class="ckb" checked="checked" />知識庫</label>
			<label for="知識家" class="ckb"><input name="FromKnowledgeHome" value="1" type="checkbox" class="ckb" checked="checked" />知識家</label>
			<label for="主題館" class="ckb"><input name="FromTopic" value="1" type="checkbox" class="ckb"  checked="checked" />主題館</label>
			<p style="clear:both;margin-left:10px;">
		 	<input name="search" type="image" src="/xslgip/style1/images3/searchBtn.gif" alt="search" class="btn" onClick="javascript:checkSearchForm(0)"/>
		  <input name="search" type="image" src="/xslgip/style1/images3/SearchBtn2.gif" alt="search" class="btn" onClick="javascript:checkSearchForm(1)"/>
		  </p>
		</form>
		<script language="javascript">
			function checkSearchForm(value)
			{				
				if( value == 0 ) {
					if( document.SearchForm.Keyword.value == "" ) {
						alert('請輸入查詢值');
						event.returnValue = false;
					}
					else {
						document.SearchForm.action = "/kp.asp?xdURL=Search/SearchResultList.asp<xsl:text disable-output-escaping="yes">&amp;</xsl:text>mp=1";
						document.SearchForm.submit();
					}
				}
				else {
					document.SearchForm.action = "/kp.asp?xdURL=Search/AdvancedSearch.asp<xsl:text disable-output-escaping="yes">&amp;</xsl:text>mp=1";
					document.SearchForm.submit();
				}
			}
		</script>
	</div>
</xsl:template>
<!--搜尋服務 End-->



<!--頁尾資訊 Start-->
<xsl:template match="hpMain" mode="xsFooter">
	<xsl:choose>
		<xsl:when test="//Footer/Article/xImgFile">
			<img>
				<xsl:attribute name="src"><xsl:value-of select="//Footer/Article/xImgFile" /></xsl:attribute>
				<xsl:attribute name="alt"><xsl:value-of select="//Footer/Article/Caption" /></xsl:attribute>
			</img>
		</xsl:when>
		<xsl:otherwise>
			<img>
				<xsl:attribute name="src">../xslgip/<xsl:value-of select="//myStyle" />/images<xsl:value-of select="//MpStyle/Article/Caption"/>/webindoor.gif</xsl:attribute>
				<xsl:attribute name="alt"><xsl:value-of select="//Footer/Article/Caption" /></xsl:attribute>
			</img>
		</xsl:otherwise>
	</xsl:choose>
	<p><xsl:value-of disable-output-escaping="yes"  select="//Footer/Article/ContentAll" /></p>
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
	content.asp?mp=<xsl:value-of select="//mp" />&amp;CuItem=<xsl:value-of select="//SitemapTop /Article/@iCuItem" />
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
