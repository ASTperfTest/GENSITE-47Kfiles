<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />


<!-- 中央主資料區 Start-->
	<xsl:template match="MainArticle">
		<!--  <div class="News"> -->		
		<h4><xsl:value-of select="Caption" /></h4>
		<xsl:if test="@IsPic ='Y'">							
			<xsl:if test="xImgFile">
				<img class="leftimg">
					<xsl:attribute name="src">
						<xsl:value-of select="xImgFile" />
					</xsl:attribute>
					<xsl:attribute name="alt">
						<xsl:value-of select="Caption" />
					</xsl:attribute>
				</img>
			</xsl:if>
		</xsl:if>
			<xsl:if test="@IsPostDate ='Y'">
				<p><em>日期:&amp;nbsp;<xsl:value-of select="PostDate" /></em></p>
			</xsl:if>
    <SPAN class="idx2" id="article" name="article">
      <xsl:value-of disable-output-escaping="yes" select="Content" />
    </SPAN>
		<!--</div>-->
	</xsl:template>


<!-- 中央主資料區 End-->


</xsl:stylesheet>
