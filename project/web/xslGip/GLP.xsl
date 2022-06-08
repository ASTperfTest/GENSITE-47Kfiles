<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:user="urn:user-namespace-here" version="1.0">

<!-- 條列／開始 --> 		
<xsl:template match="TopicList">
	<div class="album">
		<ul class="album">
			<xsl:for-each select="Article">
				<li>
					<a>
						<xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="ArticleField/xURL" /></xsl:attribute>
						<xsl:attribute name="title"><xsl:value-of select="ArticleField[fieldName='stitle']/Value" /></xsl:attribute> 
						<xsl:if test="ArticleField/xURL[@newWindow='Y']">
							<xsl:attribute name="target">_blank</xsl:attribute>
						</xsl:if>
						<xsl:value-of select="ArticleField[fieldName='stitle']/Value" />	
						<img>
							<xsl:attribute name="src">
								public/Data/<xsl:value-of disable-output-escaping="yes" select="ArticleField[fieldName='ximgFile']/Value" />
							</xsl:attribute>
							<xsl:attribute name="alt">
								<xsl:value-of disable-output-escaping="yes" select="ArticleField[fieldName='stitle']/Value" />
							</xsl:attribute>		
							<xsl:attribute name="title">
								<xsl:value-of disable-output-escaping="yes" select="ArticleField[fieldName='stitle']/Value" />
							</xsl:attribute>
						</img>
					</a>
					<xsl:value-of disable-output-escaping="yes" select="ArticleField[fieldName='xabstract']/Value" />
				</li>
			</xsl:for-each>
		</ul>
		<!--div class="more"><a href="#">更多訊息</a></div-->
	</div>
</xsl:template>
<!-- 條列／結束 --> 
</xsl:stylesheet>

