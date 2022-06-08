<?xml version="1.0"  encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="urn:user-namespace-here" version="1.0" >
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone= "yes"/>
<xsl:template match="/">
    <xsl:element name="fieldList">
	<xsl:apply-templates select="//field">
	    <xsl:sort select="fieldSeq" data-type="number" order="ascending"/>  		
	</xsl:apply-templates>
    </xsl:element>  		
</xsl:template>

<xsl:template match="field">
    <xsl:element name="field">
	<xsl:for-each select="*">
  	    <xsl:copy-of select="." />
	</xsl:for-each>  		
    </xsl:element>  
</xsl:template>

</xsl:stylesheet>