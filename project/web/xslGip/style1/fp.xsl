<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
<xsl:include href="info.xsl"/>

<xsl:template match="hpMain">
	<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
		<head>
			<xsl:apply-templates select="." mode="Header"/>
		</head>
		
		<body>
			<!--header star (style A)-->
			<xsl:apply-templates select="." mode="RunActive"/>
			<!--header End-->
			<!--nav star (style A)-->
			<xsl:apply-templates select="." mode="toplink"/>
			<!--nav end-->

			<!--layout Table star (style A)-->
			<table class="layout">
			  <tr>
				<td class="center">
					<xsl:apply-templates select="." mode="c"/>
					<!--path star (style B)-->
					<xsl:apply-templates select="." mode="xsPath"/>
					<!--path end-->
					<h3><xsl:value-of select="//xPath/UnitName"/></h3>
			
					<!-- 最新發問TAB LIST／開始 -->
					<div id="Magazine">
						<div class="Event">
							<!-- 內容區 Start -->
							<xsl:apply-templates select="//MainArticle" />
							<!-- 內容區 結束 -->
						</div>
					</div>
					<!-- 最新發問TAB LIST／結束 -->
					<!--hotessay End-->    
				</td>
			  </tr>
			</table>
			<!--layout Table End-->
			<!--footer star (style A)-->
			<div class="footer">
				<xsl:apply-templates select="." mode="xsFooter"/>		
				<xsl:apply-templates select="." mode="xsCopyright"/>		
			</div>
			<!--footer End-->
		</body>
	</html>
</xsl:template>

<!-- 中央主資料區 Start-->
<xsl:template match="MainArticle">
	<div id="cp">
		<div class="experts">
			<h4><xsl:value-of select="Caption" /></h4>
			<xsl:if test="xImgFile">
				<div class="exphoto">
					<div class="phoimg">
						<img>
							<xsl:attribute name="src"><xsl:value-of select="xImgFile" /></xsl:attribute>
							<xsl:attribute name="alt"><xsl:value-of select="Caption" /></xsl:attribute>
						</img>
					</div>
				</div>
			</xsl:if>
			<p><xsl:value-of disable-output-escaping="yes" select="Content" /></p>	
		</div>
	</div>
</xsl:template>	
</xsl:stylesheet>
