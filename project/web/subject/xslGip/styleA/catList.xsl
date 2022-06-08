<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">	
<!-- 資料大類 -->
	<xsl:template match="CatList">
		<div class="category">
  		  <a><xsl:attribute name="href">lp.asp?<xsl:value-of select="../xqURL" /></xsl:attribute>All</a>  | 
    	    <xsl:for-each select="CatItem">
  				<a>
					<xsl:attribute name="href">lp.asp?<xsl:value-of select="../../xqURL" /><xsl:value-of select="xqCondition" /></xsl:attribute>
					<xsl:value-of select="CatName" />
				</a>  | 
    	    </xsl:for-each>
	</div>
	</xsl:template>
<!-- 資料大類 -->
</xsl:stylesheet>






