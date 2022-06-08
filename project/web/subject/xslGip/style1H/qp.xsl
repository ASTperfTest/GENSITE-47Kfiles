<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
	<xsl:include href="headpart.xsl" />
	<xsl:include href="hySearch.xsl" />
	<xsl:include href="epaper.xsl"/>
	<xsl:include href="footer.xsl"/>
	<xsl:include href="menuitem.xsl"/>
	<xsl:include href="xpath.xsl"/>
	<xsl:include href="info.xsl" />
	
	<xsl:template match="hpMain">
		<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
			<xsl:apply-templates select="." mode="WInfo" />
			<head>
				<!-- 網頁標題 Start-->
				<xsl:apply-templates select="." mode="headinfo" />
				<!-- 網頁標題 End-->
			</head>
			<body class="body">
				<div id="middle">
		<table class="layout" summary="排版表格">	
			<tr>
			<!-- leftcol start -->	
				<td class="leftbg">
				<div id="leftcol"> 
				<!-- accesskey start -->
				<div class="accesskey"><a href="accesskey.htm" title="左方區塊" accesskey="L">:::</a></div>
				<!-- accesskey start -->
				
				<div class="menu">						
				<!-- 主要選單 Start-->
					<xsl:for-each select="MenuBar">
						<xsl:apply-templates select="MenuCat" />
					</xsl:for-each>
				<!-- 主要選單 End-->	
				</div>	
					
				</div>
				</td>
				<td id="center">
					<!-- accesskey start -->
					<div class="accesskey"><a href="accesskey.htm" title="中央區塊" accesskey="C">:::</a></div>
					<!-- accesskey start -->	

					<!-- friendly Start -->
					<div class="path">
						<p><xsl:apply-templates select="//xPath" /></p>
					</div>
					
					<!--主題單元查詢 Start-->					
						<xsl:apply-templates select="pHTML" />					
						<div class="top"><a><xsl:attribute name="href">dp.asp?mp=<xsl:value-of select="mp" /></xsl:attribute>首頁</a></div>					
					<!--主題單元查詢 End-->
				</td>
			</tr>
		</table>
	</div>
		
		<!-- 頁尾資訊 Start-->
				<xsl:apply-templates select="." mode="Copyright" />
		<!-- 頁尾資訊 End-->
			</body>
		</html>
	</xsl:template>

<!-- 主資料區-->
	<xsl:template match="pHTML">
			<xsl:copy-of select="." />
	</xsl:template>
<!-- 主資料區-->

</xsl:stylesheet>
