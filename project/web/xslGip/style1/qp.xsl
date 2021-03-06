<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
	<xsl:include href="headpart.xsl" />
	<xsl:include href="hySearch.xsl" />
	<xsl:include href="epaper.xsl"/>
	<xsl:include href="footer.xsl"/>

	<xsl:template match="hpMain">
		<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
			<head>
				<title>GIP</title>
				<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
				<link href="xslGIP/style1/css/home.css" rel="stylesheet" type="text/css" />
			</head>
			<body>
		<!-- 網頁標題 Start-->
					<xsl:apply-templates select="." mode="headinfo" />
		<!-- 網頁標題 End-->

		<!-- 快速導覽列 Start-->
				<div class="Title"></div>
					<xsl:apply-templates select="MenuBar2" />
		<!-- 快速導覽列 End-->

		<!-- 左方功能選單區 Start-->
				<div class="Func">
				<!-- 主要選單 Start-->
					<xsl:for-each select="MenuBar">
						<xsl:apply-templates select="MenuCat" />
					</xsl:for-each>
				<!-- 主要選單 End-->
				<!-- AD Start-->
					<xsl:apply-templates select="AD" />
				<!-- AD End-->
				</div> <!--class="Func"-->
		<!-- 左方功能選單區 End-->

		<!--主題單元查詢 Start-->
				<div class="Main">
					<div class="News">
						<div class="Path">
							<xsl:apply-templates select="//xPath" />
						</div>
						<div class="Head">
							<xsl:value-of select="//xPath/UnitName" />
						</div>
						<xsl:apply-templates select="pHTML" />
					</div>
				</div>
		<!--主題單元查詢 End-->

		<!-- 右方資料區 Start-->
				<div class="Column">
					<xsl:apply-templates select="." mode="xsSearch" />
					<xsl:apply-templates select="." mode="xsEPaper" />
				</div>
		<!-- 右方資料區 End-->
		
		<!-- 頁尾資訊 Start-->
				<xsl:apply-templates select="Copyright" />
		<!-- 頁尾資訊 End-->
			</body>
		</html>
	</xsl:template>
	
<!-- 快速導覽列 -->
	<xsl:template match="MenuBar2">
		<div class="Quickview">
		<xsl:for-each select="MenuCat">
			│<a>
					<xsl:attribute name="href">
						<xsl:value-of select="redirectURL" />
					</xsl:attribute>
					<xsl:if test="@newWindow='Y'">
						<xsl:attribute name="target">_nwGip</xsl:attribute>
					</xsl:if>
					<xsl:value-of select="Caption" />
				</a>		
		</xsl:for-each>
	│</div>
	</xsl:template>
<!-- 快速導覽列 -->

<!-- 主要選單 -->
	<xsl:template match="MenuCat">
		<div class="Link">
			<div class="Head">
				<a>
					<xsl:attribute name="href">
						<xsl:value-of select="redirectURL" />
					</xsl:attribute>
					<xsl:if test="@newWindow='Y'">
						<xsl:attribute name="target">_nwGip</xsl:attribute>
					</xsl:if>
					<xsl:value-of select="Caption" />
				</a>
			</div>
			<xsl:if test="MenuItem">
				<div class="Body">
					<xsl:for-each select="MenuItem">
						<xsl:apply-templates select="." />
					</xsl:for-each>
				</div>
			</xsl:if>
		</div>
	</xsl:template>

	<xsl:template match="MenuItem">
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="redirectURL" />
			</xsl:attribute>
			<xsl:if test="@newWindow='Y'">
				<xsl:attribute name="target">_nwGip</xsl:attribute>
			</xsl:if>
			<xsl:value-of select="Caption" />
		</a>
	</xsl:template>
<!-- 主要選單 -->

<!--  路徑連結  -->
	<xsl:template match="xPath">
		<a href="mp.asp">首頁</a>
		<xsl:for-each select="xPathNode">
			<img src="xslGIP/style1/images/arrow_path.gif" alt=" " />
			<a>
				<xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="@xNode" /></xsl:attribute>
				<xsl:value-of select="@Title" />
			</a>
		</xsl:for-each>
		<!--xsl:value-of select="//xPath/UnitName" / -->
	</xsl:template>
<!--  路徑連結  -->

<!-- 主資料區-->
	<xsl:template match="pHTML">
			<xsl:copy-of select="." />
	</xsl:template>
<!-- 主資料區-->

<!-- AD Start -->
	<xsl:template match="AD">
		<div class="AD">
			<p>
				<xsl:for-each select="//AD/Article">
					<a>
						<xsl:attribute name="href">
							<xsl:value-of select="xURL" />
						</xsl:attribute>
						<xsl:if test="@newWindow='Y'">
							<xsl:attribute name="target">_nwGip</xsl:attribute>
						</xsl:if>
						<img border="0" width="120">
							<xsl:attribute name="src">
								<xsl:value-of select="xImgFile" />
							</xsl:attribute>
							<xsl:attribute name="alt">
								<xsl:value-of select="Caption" />
							</xsl:attribute>
						</img>
					</a>
				</xsl:for-each>
			</p>
		</div>
	</xsl:template>
<!-- AD End-->

</xsl:stylesheet>
