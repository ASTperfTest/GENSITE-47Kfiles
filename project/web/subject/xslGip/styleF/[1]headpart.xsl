<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">	
	<xsl:include href="hySearch.xsl" />
	<xsl:template match="hpMain" mode="headinfo">
		<!-- header start -->
			<div id="header"> 
					<script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
					</script>
					<script type="text/javascript">
					_uacct = "UA-1667518-10";
					urchinTracker();
					</script>
					<script src="/nocopy.js" language="javascript"></script>
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






