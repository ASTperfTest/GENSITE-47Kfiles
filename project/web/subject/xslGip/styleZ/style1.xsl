<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />




<!-- 主資料區-->
	<xsl:template match="TopicList">
		<xsl:value-of disable-output-escaping="yes" select="../HeaderPart" />		
				<caption></caption>
					<tr>
						<th width="8%">序號</th>
						<xsl:if test="Article/@IsPostDate='Y'" >
						<th width="75%">標題</th>						
						<th width="17%">發佈日期</th>
						</xsl:if>
						<xsl:if test="Article/@IsPostDate='N'" >
							<th width="92%">標題</th>
						</xsl:if>
					</tr>
				
				<xsl:for-each select="Article">
				<tr>					
					<xsl:apply-templates select="." />
				</tr>
				</xsl:for-each>				
		<xsl:value-of disable-output-escaping="yes" select="../FooterPart" />
	</xsl:template>

	<xsl:template match="TopicList/Article">
		
			<td>
				<xsl:value-of select="NoNum" />			
			</td>
			<td>
			<xsl:if test="xURL !=''">
				<a>
					<xsl:attribute name="href">
						<xsl:value-of select="xURL" />
					</xsl:attribute>
					<xsl:attribute name="title">
						<xsl:value-of select="Caption" />
					</xsl:attribute>
					<xsl:if test="@newWindow='Y'">
						<xsl:attribute name="target">_nwGip</xsl:attribute>
					</xsl:if>
					<xsl:value-of select="Caption" />
				</a>
			</xsl:if>
			<xsl:if test="xURL =''">
				<xsl:value-of select="Caption" />
			</xsl:if>
			</td>			
			<td>
			<xsl:if test="@IsPostDate='Y'">
				<xsl:value-of select="PostDate" />
			</xsl:if>
			</td>
		
			
	</xsl:template>
<!-- 主資料區-->



</xsl:stylesheet>
