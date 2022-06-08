<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
	

	<xsl:template match="MainArticle">
	<div>
		<h4><xsl:value-of select="MainArticleField[fieldName='stitle']/Value" /></h4>	
		<SPAN class="idx2" id="article" name="article">		
		<xsl:element name="iframe">
			<xsl:attribute name="scrolling">no</xsl:attribute>
			<xsl:attribute name="width">95%</xsl:attribute>   
			<xsl:attribute name="border">0</xsl:attribute>	
			<xsl:attribute name="frameborder">0</xsl:attribute>
			<xsl:attribute name="cellspacing">0</xsl:attribute>			
			<xsl:attribute name="height">450</xsl:attribute>
			<xsl:attribute name="src">/subject/FileShowImage.aspx<xsl:value-of select="//qStr" /></xsl:attribute>

		</xsl:element>
		</SPAN>
	</div>
	</xsl:template>
	
	<!--取消附件顯示-->
	<xsl:template match="AttachmentList" mode="listx">
	</xsl:template>	

</xsl:stylesheet>
