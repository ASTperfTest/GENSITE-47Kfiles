<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
<xsl:include href="info.xsl" />

<xsl:template match="hpMain">
<html>
	<head>
		<xsl:apply-templates select="." mode="topinfo"/>		
		<link rel="stylesheet" type="text/css" >
			<xsl:attribute name="href">/xslgip/<xsl:value-of select="//myStyle"/>/css/print.css</xsl:attribute>
		</link>		
	</head>

	<body>
		<h1><xsl:value-of select="$title"/></h1>
		
		<!-- 瀏覽路徑／開始 -->
		<xsl:apply-templates select="." mode="xsPath"/>
		<!-- 瀏覽路徑／結束 -->
		
		<!-- 文章／開始 -->
			<xsl:apply-templates select="MainArticle"/>
		<!-- 文章／結束 -->
		
		<!-- 相關聯結／開始 -->
		<xsl:if test="ReferenceList">
			<xsl:apply-templates select="ReferenceList"/>
		</xsl:if>
		<!-- 相關聯結／結束 -->
	</body>
</html>
</xsl:template>
</xsl:stylesheet>
