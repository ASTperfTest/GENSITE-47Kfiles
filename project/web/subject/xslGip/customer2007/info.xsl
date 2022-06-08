<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	
<xsl:variable name="title">農業知識入口網</xsl:variable>	
	

<!-- 網站標題／開始 -->
<xsl:template match="hpMain" mode="webinfo2">
	<xsl:apply-templates select="." mode="u"/>
	<div class="WebTitle">
		<img>
			<xsl:attribute name="src">/xslgip/<xsl:value-of select="//myStyle"/>/images/webtitle.jpg</xsl:attribute>
			<xsl:attribute name="alt"><xsl:value-of select="$title"/></xsl:attribute>
		</img>
	</div>
</xsl:template>

<xsl:template match="hpMain" mode="webinfo">
	<xsl:apply-templates select="." mode="u"/>
	<div class="WebTitle">
		<img>
			<xsl:attribute name="src">/xslgip/<xsl:value-of select="//myStyle"/>/images/webtitle2.jpg</xsl:attribute>
			<xsl:attribute name="alt"><xsl:value-of select="$title"/></xsl:attribute>
		</img>
	</div>
</xsl:template>
<!-- 網站標題／結束 -->



<!-- 導覽列／開始 -->
<xsl:template match="hpMain" mode="info">
	<div class="Nav"> 
		<ul> 
			<xsl:apply-templates select="." mode="submenu"/>		
		</ul>
	</div> 
</xsl:template>
<!-- 導覽列／結束 -->



<!-- 站內查詢／開始 -->
<xsl:template match="hpMain" mode="Search">
	<xsl:apply-templates select="." mode="searchC3"/>
</xsl:template>
<!-- 站內查詢／結束 -->



<!-- 訂閱電子報／開始 -->
<xsl:template match="hpMain" mode="xsEPaper">
	<xsl:apply-templates select="." mode="epaperC3"/>	
</xsl:template>
<!-- 訂閱電子報／結束 -->
</xsl:stylesheet>
