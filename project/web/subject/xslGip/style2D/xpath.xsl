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
		<!--xsl:value-of select="//xPath/UnitName" / -->
	</xsl:template>
<!--  路徑連結 End -->
</xsl:stylesheet>






