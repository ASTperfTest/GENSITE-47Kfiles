<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:hyweb="urn:gip-hyweb-com" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
<xsl:include href="commoninfo.xsl"/>

<xsl:template match="hpMain">
	<html xmlns="http://www.w3.org/1999/xhtml">
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
			<!--subnav star (style G)-->
			<xsl:apply-templates select="." mode="weblink"/>
			<!--subnav end-->
			<!--layout Table star (style A)-->
			<table class="layout">
			  <tr>
				<td class="left">
					<xsl:apply-templates select="." mode="l"/>
					<!--season star-->
					<xsl:apply-templates select="." mode="SolarTerms"/>
					<!--season end-->
					<!--menu star (style C)-->
					<xsl:apply-templates select="." mode="Menu"/>
					<!--menu end-->
			
					<!--RSS star-->
					<xsl:apply-templates select="." mode="xsRSS"/>
					<!--RSS end-->
					<!--ad star style A-->
					<xsl:apply-templates select="." mode="xsAD"/>
					<!--ad end-->
				</td>
					<!--path star (style B)-->
					<!--xsl:apply-templates select="." mode="xsPath"/-->
					<!--path end-->
					<!--h3><xsl:value-of select="//xPath/UnitName"/></h3-->
			
					<!-- 最新發問TAB LIST／開始 -->
					<!--div id="Magazine">
						<div class="Event"-->
							<!-- 內容區 Start -->
							<xsl:apply-templates select="pHTML"/>
							<!-- 內容區 結束 -->
						<!--/div>
					</div-->
					<!-- 最新發問TAB LIST／結束 -->
					<!--hotessay End-->    

			  </tr>
			</table>
			<!--layout Table star (style A)-->		
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
<xsl:template match="pHTML">
<xsl:value-of disable-output-escaping="yes" select="." />
<!--xsl:copy-of select="."/-->
</xsl:template>
</xsl:stylesheet>
