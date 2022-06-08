<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
<xsl:include href="info.xsl"/>

<xsl:template match="hpMain">
	<html xmlns="http://www.w3.org/1999/xhtml">
		<head>
			<xsl:apply-templates select="." mode="Header"/>
			<link rel="stylesheet" type="text/css">
					<xsl:attribute name="href">/js/ui.datepicker.css</xsl:attribute>
					<xsl:attribute name="charset">utf-8</xsl:attribute>
				</link>
				<script src="/js/jquery.js" language="javascript"  ></script>
				<script src="/js/ui.datepicker.js" language="javascript" ></script>
				<script src="/js/ui.datepicker-zh-TW.js" language="javascript"  ></script>
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
			<xsl:apply-templates select="pHTML"/>
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
