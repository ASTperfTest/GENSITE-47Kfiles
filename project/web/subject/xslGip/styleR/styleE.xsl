<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />

<xsl:template match="MainArticle">
		<div id="cp5">
        <h4><xsl:value-of select="MainArticleField[fieldName='stitle']/Value" /></h4>		
		
					<xsl:if test="MainArticleField[fieldName='ximgFile']/Value!=''">
						<div class="imageY">
							<img>
								<xsl:attribute name="src">/public/Data/<xsl:value-of select="MainArticleField[fieldName='ximgFile']/Value" /></xsl:attribute>
								<xsl:attribute name="alt"><xsl:value-of select="MainArticleField[fieldName='stitle']/Value" /></xsl:attribute>
							</img>
						</div>				
					</xsl:if>
					<p>
            <SPAN class="idx2" id="article" name="article">
              <xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xbody']/Value" />
            </SPAN>
          </p>	
				</div>			
				<!--xsl:apply-templates select="AttachmentList" mode="listx"/>					
				<xsl:apply-templates select="//AttachmentList" /-->
</xsl:template>

	<xsl:template match="AttachmentList" mode="listx">
			<h4>附件</h4>
				<div class="yoxview">
				<xsl:for-each select="Attachment">
						<div class="attachmentX">							
							<a>
								<xsl:attribute name="href">
									<xsl:value-of select="URL" /> 
								</xsl:attribute>
								<xsl:attribute name="title">
									<xsl:value-of select="Caption" /> 
								</xsl:attribute>		  
								<img class="imageX">
									<xsl:attribute name="src">
										<xsl:value-of select="URL" /> 
									</xsl:attribute>
									<xsl:attribute name="alt">
										<xsl:value-of select="Caption" /> 
									</xsl:attribute>
									<xsl:attribute name="title">
										<xsl:value-of select="Caption" /> 
									</xsl:attribute>
								</img>
								<xsl:value-of select="Caption" />
							</a>
						<p><xsl:value-of select="Descxx" /></p>
						</div>
				</xsl:for-each>
				</div>
	</xsl:template>

</xsl:stylesheet>
