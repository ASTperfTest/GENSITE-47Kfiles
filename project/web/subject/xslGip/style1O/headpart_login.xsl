<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">	
	<xsl:include href="hySearch.xsl" />
	<xsl:include href="member.xsl" />
	<xsl:template match="hpMain" mode="headinfo">
		<!-- header start -->
		  <link rel="stylesheet" type="text/css">
		  	<xsl:attribute name="href">xslgip/<xsl:value-of select="//MpStyle" />/css/login.css</xsl:attribute>
      </link >
		 
		  
			<div id="login" ><xsl:apply-templates select="." mode="xsMember" /></div>
      <div id="header">
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
					<xsl:if test="//lockrightbtn='Y'"><script src="/nocopy.js" language="javascript" /></xsl:if>

			<!-- accesskey start -->
			<div class="accesskey">
				<a href="accesskey.htm" title="上方區塊" accesskey="U">:::</a>
			</div>
			<!-- accesskey start -->
			
				<span class="logo">
					<a>
						<xsl:attribute name="href">dp.asp?mp=<xsl:value-of select="mp" /></xsl:attribute>
						<img>
							<xsl:attribute name="src">
								<xsl:value-of select="TitlePicFilePath" />
							</xsl:attribute>
							<xsl:attribute name="alt">
								<xsl:value-of select="TitlePicFileName" />
							</xsl:attribute>					
						</img>
					</a>
				</span>
				
				<ul>
					<li><a>
						<xsl:attribute name="href">
							dp.asp?mp=<xsl:value-of select="mp" />
						</xsl:attribute>回首頁
					</a></li>
					<!--li><a><xsl:attribute name="href">mailto:<xsl:value-of select="contactmail" /></xsl:attribute>聯絡我們</a></li-->
					<li><a><xsl:attribute name="href">sitemap.asp?mp=<xsl:value-of select ="mp" /></xsl:attribute>網站導覽</a></li>
				</ul>
				<xsl:apply-templates select="." mode="xsSearch" />
				
			</div>
			<!--橫幅圖片 start -->
			<div class="image">
					<img>
						<xsl:attribute name="src">
							<xsl:value-of select="BannerPicFilePath" />
						</xsl:attribute>
						<xsl:attribute name="alt">
							<xsl:value-of select="BannerPicFileName" />
						</xsl:attribute>
					</img>
			</div>
			<!--橫幅圖片 end-->
	</xsl:template>
</xsl:stylesheet>






