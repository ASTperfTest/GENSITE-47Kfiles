<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />


<xsl:template match="MainArticle">
		<div id="cp6">
				<h4><xsl:value-of select="MainArticleField[fieldName='stitle']/Value" /></h4>		
				<div class="column">					
					<xsl:if test="MainArticleField[fieldName='ximgFile']/Value!=''">
							<img class="imageW">
								<xsl:attribute name="src">/public/Data/<xsl:value-of select="MainArticleField[fieldName='ximgFile']/Value" /></xsl:attribute>
								<xsl:attribute name="alt"><xsl:value-of select="MainArticleField[fieldName='stitle']/Value" /></xsl:attribute>
							</img>
					</xsl:if>
					<p>
            <SPAN class="idx2" id="article" name="article">
              <xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xbody']/Value" />
            </SPAN>
          </p>						
				</div>
				<div class="thumb">	
				<div class="yoxview">				
				<xsl:for-each select="//AttachmentList/Attachment">
					<div class="imageBlock">
						<xsl:if test="fileType='jpg' or fileType='gif'">					
							<div class="imageS">
								<a>
									<xsl:attribute name="href">
										<xsl:value-of select="URL" /> 
									</xsl:attribute>
									<xsl:attribute name="title">
										<xsl:value-of select="Caption" /> 
									</xsl:attribute>		  
								<img>
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
								</a>
							</div>
								<xsl:choose>
									<xsl:when test="Descxx!=''">
										<p>
											<xsl:value-of select="Descxx" />
										</p>
									</xsl:when>
									<xsl:otherwise>
										<p>
											<xsl:value-of select="Caption" />
										</p>
									</xsl:otherwise>
								</xsl:choose>										
						</xsl:if>				
					</div>
				</xsl:for-each>
				</div>	
				</div>				
				<!--xsl:apply-templates select="//AttachmentList" /-->
		</div>
	</xsl:template>

	<xsl:template match="AttachmentList" mode="listx">
			<h4>附件</h4>
			<ul>
				<xsl:for-each select="Attachment">
					<xsl:if test="not(fileType='jpg') and not(fileType='gif')">					
					<li>					
						<div class="title">
							<a target="_nwGIP">
								<xsl:attribute name="href"><xsl:value-of select="URL" /></xsl:attribute>
								<xsl:value-of select="Caption" />
							</a>
						</div>									
					</li>
					</xsl:if>				
				</xsl:for-each>
			</ul>
	</xsl:template>	

</xsl:stylesheet>
