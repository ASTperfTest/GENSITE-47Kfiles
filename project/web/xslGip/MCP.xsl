<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:user="urn:user-namespace-here" version="1.0">
<xsl:template match="MainArticle">
	<div id="cp">
		<div class="experts">
			<h4><xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='stitle']/Value" /></h4>
			<div class="video">
				<embed width="300" height="240" autostart="ture">
				<xsl:attribute name="src"><xsl:value-of disable-output-escaping="yes" select="//AttachmentList/Attachment/URL"/></xsl:attribute>
				</embed>
			</div>
			<table summary="排版表格" class="cptable">
				<col class="title" />
				<col class="cptablecol2" />
				<xsl:for-each select="MainArticleField">
					<tr>
						<xsl:apply-templates select="."/>
					</tr>
				</xsl:for-each>
			</table>
			<!--p><xsl:value-of disable-output-escaping="yes" select="//MainArticleField[fieldName='xbody']/Value" /></p-->
		</div>
	</div>
	
	<!--div id="cp">
		<div class="experts">
			<h4><xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='stitle']/Value" /></h4>
			<div class="video">
				<embed width="300" height="240" autostart="false">
				<xsl:attribute name="src">public/MMO/video/<xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='mmoFileName']/Value"/></xsl:attribute>
				</embed>
			</div>
			<p><xsl:value-of disable-output-escaping="yes" select="//MainArticleField[fieldName='xbody']/Value" /></p>
		</div>
	</div-->
</xsl:template>

<xsl:template match="MainArticleField">    
	<th scope="row"><xsl:value-of select="Title" /></th>
	<xsl:choose>
		<xsl:when test="fieldName='ximgFile'">
			<td>
				<img>
					<xsl:attribute name="src">
						public/Data/<xsl:value-of disable-output-escaping="yes" select="../MainArticleField[fieldName='ximgFile']/Value" />
					</xsl:attribute>
					<xsl:attribute name="alt">
						<xsl:value-of disable-output-escaping="yes" select="../MainArticleField[fieldName='stitle']/Value" />
					</xsl:attribute>		
					<xsl:attribute name="title">
						<xsl:value-of disable-output-escaping="yes" select="../MainArticleField[fieldName='stitle']/Value" />
					</xsl:attribute>
				</img>
			</td>
		</xsl:when>
		<xsl:when test="fieldName='xurl'">
			<td>
				<a>
					<xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="../MainArticleField[fieldName='xurl']/Value" /></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="Value" /></xsl:attribute> 
					<xsl:if test="xURL[@newWindow='Y']">
						<xsl:attribute name="target">_blank</xsl:attribute>
					</xsl:if>
					<xsl:value-of select="Value" />       	        		
				</a>			
			</td>
		</xsl:when>
		
		<xsl:otherwise><td><xsl:value-of disable-output-escaping="yes" select="Value" /></td></xsl:otherwise>
	</xsl:choose>
</xsl:template>   
</xsl:stylesheet>

