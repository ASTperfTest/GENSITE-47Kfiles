<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:user="urn:user-namespace-here" version="1.0">

<xsl:template match="TopicList">
	<div class="lp">
		<xsl:value-of disable-output-escaping="yes" select="../HeaderPart" />
		<!--標准文章樣式-->
		<table width="100%" border="0" cellspacing="0" cellpadding="0" summary="問題列表資料表格">
		<xsl:apply-templates select="../TopicTitleList" />		
				<xsl:for-each select="Article">
					<tr>
						<td><xsl:value-of select="position()"/></td>
						<xsl:for-each select="ArticleField">
							<xsl:apply-templates select="."/>
						</xsl:for-each>
					</tr>
				</xsl:for-each>
		</table>
		<xsl:value-of disable-output-escaping="yes" select="../FooterPart" />
	</div>
</xsl:template>
    
<xsl:template match="TopicTitleList">    
  	<tr>
		<th scope="col">&amp;nbsp;</th>
		<xsl:for-each select="TopicTitle"> 			
			<th scope="col"><xsl:value-of select="." /></th>
		</xsl:for-each>			
	</tr>
</xsl:template>

<xsl:template match="ArticleField">    
	<xsl:choose>	
		<xsl:when test="xURL">
			<td>
				<a>
					<xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="Value" /></xsl:attribute> 
					<xsl:if test="xURL[@newWindow='Y']">
						<xsl:attribute name="target">_blank</xsl:attribute>
					</xsl:if>
					<xsl:value-of select="Value" />       	        		
				</a>			
			</td>
		</xsl:when>
		<xsl:when test="Value=''"><td>&amp;nbsp;</td></xsl:when>
		<xsl:otherwise><td><xsl:value-of select="Value" /></td></xsl:otherwise>
	</xsl:choose>	
</xsl:template>
</xsl:stylesheet>

