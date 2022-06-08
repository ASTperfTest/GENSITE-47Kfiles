<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
	<xsl:include href="hySearch.xsl" />
	<xsl:include href="AttachmentList.xsl" />
	<xsl:include href="info.xsl" />
	<xsl:include href="footer.xsl"/>
	<xsl:template match="hpMain">
		<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
			<xsl:apply-templates select="." mode="WInfo" />

			<body>
				<xsl:apply-templates select="MainArticle" />

				<!-- 頁尾資訊 Start-->
				<xsl:apply-templates select="." mode = "Copyright"/>
				<!-- 頁尾資訊 End-->
			</body>
		</html>
	</xsl:template>

	<!-- 中央主資料區 Start-->
	<xsl:template match="MainArticle">
		<xsl:if test="//lockrightbtn='Y'">
			<script src="/nocopy.js" language="javascript"></script>
		</xsl:if>
		<!--  <div class="News"> -->
		<div id="center">
			<h4>
				<xsl:value-of select="Caption" />
			</h4>
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
				<p>
					<em>
						<xsl:value-of select="PostDate" />
					</em>
				</p>
			</xsl:if>
			<xsl:value-of disable-output-escaping="yes" select="Content" />
		</div>
		<!--</div>-->
	</xsl:template>
	<!-- 中央主資料區 End-->



	<xsl:template match="RelatedList">
		<div class="RelArtical">
			<div class="Head">相關文章</div>
			<div class="Body">
				<ul>
					<xsl:for-each select="RelatedRef">
						<xsl:apply-templates select="." />
					</xsl:for-each>
				</ul>
			</div>
		</div>
	</xsl:template>
	<xsl:template match="RelatedRef">
		<li>
			<a>
				<xsl:attribute name="href">
					<xsl:value-of select="URL" />
				</xsl:attribute>
				<xsl:attribute name="title">
					<xsl:value-of select="keyWeights" />
				</xsl:attribute>
				<xsl:value-of select="Caption" />
			</a>
		</li>
	</xsl:template>
</xsl:stylesheet>
