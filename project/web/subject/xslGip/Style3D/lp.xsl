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
	<xsl:include href="catList.xsl" />
	<xsl:include href="pageissue.xsl" />
	<xsl:include href="info.xsl" />

	<xsl:template match="hpMain">
		<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
			<xsl:apply-templates select="." mode="WInfo" />

			<body class="body">
				<!-- 網頁標題 Start-->
				<xsl:apply-templates select="." mode="headinfo" />
				<!-- 網頁標題 End-->
				<div id="middle">
					<table class="layout" summary="排版表格">
						<tr>
							<!-- leftcol start -->
							<td class="leftbg">
								<div id="leftcol">

									<!-- accesskey start -->
									<div class="accesskey">
										<a href="accesskey.htm" title="左方區塊" accesskey="L">:::</a>
									</div>
									<!-- accesskey start -->



									<!-- 左方功能選單區 Start-->

									<!-- 主要選單 start -->
									<div class="menu">
										<ul>
											<xsl:for-each select="MenuBar">
												<xsl:apply-templates select="MenuCat" />
											</xsl:for-each>
										</ul>
									</div>
									<!-- 主要選單 end -->
								</div>
							</td>
							<!-- 左方功能選單區 End-->

							<!-- 中央主資料區 Start-->
							<td id="center">

								<!-- accesskey start -->
								<div class="accesskey">
									<a href="accesskey.htm" title="中央區塊" accesskey="C">:::</a>
								</div>
								<!-- accesskey start -->

								<!-- friendly Start -->

								<div class="Path">
									<p>
										<xsl:apply-templates select="//xPath" />
									</p>
								</div>

								<h4>
									<xsl:value-of select="//xPath/UnitName" />
								</h4>
								<!--title end-->

								<xsl:apply-templates select="CatList" />
								<!--資料大類 END-->

								<xsl:apply-templates select="." mode="pageissue" />
								<!--分頁 END-->
								<table class="lptable" summary="資料表格">
									<xsl:apply-templates select="TopicList" />
								</table>
								<!--標題列 END-->

								<!--<xsl:apply-templates select="." mode="pageissue" />-->
								<!--分頁 END-->
								<div style="text-align: right">
									<a href="#top">
										<img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方"  border="0"/>
									</a>
								</div>
								<!-- 中央主資料區 End-->
							</td>
							<!-- content end -->
						</tr>
					</table>
				</div>


				<!-- 頁尾資訊 Start-->
				<xsl:apply-templates select="." mode = "Copyright"/>
				<!-- 頁尾資訊 End-->
			</body>
		</html>
	</xsl:template>



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
