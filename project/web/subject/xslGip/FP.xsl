<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:user="urn:user-namespace-here" version="1.0">

    <xsl:template match="MainArticle">		
 	    <div id="urlnavbar"><xsl:apply-templates select="//xPath"/><!--xsl:value-of select="//xPath/UnitName" /--></div>
	   <div id="standarticle"><!--標准文章樣式-->
  <p class="contenttitle"> </p>
		<Table>
    	    	<xsl:for-each select="MainArticleField">
  		    <TR>
    	            <xsl:apply-templates select="."/>
		    </TR>
    	    	</xsl:for-each>
  		</Table>
		</div>
    </xsl:template>
    
  <xsl:template match="MainArticleField">    
		<TH nowrap="true"><xsl:value-of select="Title" /></TH>
		<TD><xsl:value-of disable-output-escaping="yes" select="Value" /></TD>
  </xsl:template>    

    <xsl:template match="AttachmentList">
		<div id="attachmenu">附件列表</div>
		<div id="attachlist">
  		<ol>
    	    <xsl:for-each select="Attachment">
    	            <xsl:apply-templates select="."/>
    	    </xsl:for-each>
 		</ol>		
	   	</div>
    </xsl:template>

    <xsl:template match="Attachment">
		<li><a target="_nwGIP">
		  <xsl:attribute name="href"><xsl:value-of select="URL" /></xsl:attribute>
		  <xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute>
		  <xsl:value-of select="Caption" /></a>
		</li>
    </xsl:template>
    

</xsl:stylesheet>

