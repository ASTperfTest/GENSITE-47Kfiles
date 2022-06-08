<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
<xsl:include href="info.xsl" />
<xsl:include href="../form.xsl" />

<xsl:template match="hpMain">
	<html>
		<head>
			<xsl:apply-templates select="." mode="topinfo"/>		
			<link rel="stylesheet" type="text/css" >
				<xsl:attribute name="href">/xslgip/<xsl:value-of select="//myStyle"/>/css/contactus.css</xsl:attribute>
			</link>				
		</head>
		
		<body>
			<xsl:apply-templates select="." mode="u"/>
		
			<!-- 網站標題／開始 -->
				<xsl:apply-templates select="." mode="webinfo"/>
			<!-- 網站標題／結束 -->
	
			<!-- 導覽列／開始 -->
				<xsl:apply-templates select="." mode="info"/>
			<!-- 導覽列／結束 -->
	
			<!-- 站內查詢／開始 -->
				<xsl:apply-templates select="." mode="Search"/>
			<!-- 站內查詢／結束 -->
		
			<!-- 排版表格／開始 -->
			<table class="Layout" summary="排版表格">
			  <tr>
					<!-- 左欄／開始 -->
					<td class="Left">
						<xsl:apply-templates select="." mode="m"/>
						<!-- 選單／開始 -->
							<xsl:apply-templates select="." mode="Menu" /> 
						<!-- 選單／結束 -->
					</td>
					<!-- 左欄／結束 -->
			
					<!-- 中欄／開始 -->
					<td class="Center">
						<xsl:apply-templates select="." mode="c"/>
			
						<!-- 單元名稱／開始 -->
						<div class="PageTitle"><xsl:value-of select="//PageTitle" /></div>
						<!-- 單元名稱／結束 -->
			
						<!-- 瀏覽路徑／開始 -->
						<xsl:apply-templates select="xPath" />
						<!-- 瀏覽路徑／結束 -->
			
						<!-- 聯絡我們表單／開始 -->
						<xsl:apply-templates select="." mode="forward" />
						<!-- 聯絡我們表單／結束 -->

					</td>
				<!-- 中欄／結束 -->  
				</tr>
			</table>
			<!-- 排版表格／結束 -->
			
			<!-- 頁尾資訊／開始 -->
				<div class="Bottom2">
					<xsl:apply-templates select="." mode="bottominfo" />
				</div>
			<!-- 頁尾資訊／結束 -->
		</body>
	</html>
</xsl:template>



<!--  路徑連結  -->
<xsl:template match="xPath">
	<div class="Path">
		<xsl:value-of select="//xxpath" />
			<a >
				<xsl:attribute name="href">dp.asp?mp=<xsl:value-of select="//mp"/></xsl:attribute>
				<xsl:attribute name="title"><xsl:value-of select="//xxhome" /></xsl:attribute><xsl:value-of select="//xxhome" />
			</a>
		<img alt="向右箭頭">
			<xsl:attribute name="src">/xslgip/<xsl:value-of select="//myStyle"/>/images/path_arrow.gif</xsl:attribute>
		</img><xsl:value-of select="//PageTitle" />
	</div>
</xsl:template>  
<!--  路徑連結  -->	

</xsl:stylesheet>