<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">	
	<!--  路徑連結 Start -->
	<xsl:template match="xPath">
		<a><xsl:attribute name="href">dp.asp?mp=<xsl:value-of select="../mp" /></xsl:attribute>首頁</a>
		<xsl:for-each select="xPathNode">
			&gt;
			<a>
				<xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="@xNode" /></xsl:attribute>
				<xsl:value-of select="@Title" />
			</a>
		</xsl:for-each>
		<xsl:for-each select="//CatList/mediapath">
			<xsl:if test="mcode!=''">
				>
				<xsl:choose>
					<xsl:when test="murl!=''">
						<a>
							<xsl:attribute name="href"><xsl:value-of select="murl" />&amp;mp=<xsl:value-of select="//mp" /></xsl:attribute>
							<xsl:attribute name="title"><xsl:value-of select="mvalue" /></xsl:attribute>
							<xsl:value-of select="mvalue" />
						</a>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="mvalue" />
					</xsl:otherwise>
				</xsl:choose>
			</xsl:if>
		</xsl:for-each>
		<!--xsl:value-of select="//xPath/UnitName" / -->
		<!--路徑套用主分類 START-->
        <xsl:for-each select="//CatList/CatItem">
		<xsl:if test="contains(concat(//qURL, ''), concat(xqCondition, ''))"> > </xsl:if>
        <xsl:if test="contains(concat(//qURL, ''), concat(xqCondition, ''))">
         <a>
        <xsl:attribute name="href">lp.asp?<xsl:value-of select="../../xqURL" /><xsl:value-of select="xqCondition" /></xsl:attribute>
        <xsl:attribute name="title"><xsl:value-of select="CatName" /></xsl:attribute>
        <xsl:value-of select="CatName" />
        </a>
        </xsl:if>

        </xsl:for-each>
        <!--路徑套用主分類 END-->
	</xsl:template>
<!--  路徑連結 End -->
</xsl:stylesheet>






