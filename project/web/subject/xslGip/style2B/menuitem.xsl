<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">	
	<xsl:template match="MenuCat">
			<!--第一層目錄 start-->
			<li>
      <a>
				<xsl:attribute name="href">
					<xsl:value-of select="redirectURL" />
				</xsl:attribute>
				<xsl:attribute name="title">
						<xsl:value-of select="CaptionTip" />
			    </xsl:attribute>
				<xsl:if test="@newWindow='Y'">
					<xsl:attribute name="target">_nwGip</xsl:attribute>
				</xsl:if>
				<xsl:value-of select="Caption" />
			</a>
			<!--第一層目錄 end-->
				<xsl:if test="MenuItem">
				<ul>
					<xsl:for-each select="MenuItem">				
						<li><xsl:apply-templates select="." /></li>
					</xsl:for-each>
				</ul>
				</xsl:if>
			</li>
	</xsl:template>
	
	<xsl:template match="MenuItem">
		<!--項目或子節點 start -->
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="redirectURL" />
			</xsl:attribute>
			<xsl:attribute name="title">
						<xsl:value-of select="CaptionTip" />
			    </xsl:attribute>
			<xsl:if test="@newWindow='Y'">
				<xsl:attribute name="target">_nwGip</xsl:attribute>
			</xsl:if>
			<xsl:value-of select="Caption" />
		</a>
		<!--項目或子節點 end-->
	</xsl:template>
</xsl:stylesheet>






