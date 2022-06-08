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

	
	<xsl:template match="hpMain">
		<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
			<head>
				<title><xsl:value-of select = "rootname" /></title>
				<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />				
				<link rel="stylesheet" type="text/css" href="xslGIP/style4/css/style32.css" />
			</head>		
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
						<div class="accesskey"><a href="accesskey.htm" title="左方區塊" accesskey="L">:::</a></div>
						<!-- accesskey start -->	
	

						<!-- leftbutton start -->
						<div class="menu">

						<!-- 左方功能選單區 Start-->
				
						<!-- 主要選單 Start-->
						<xsl:for-each select="MenuBar">
							<xsl:apply-templates select="MenuCat" />
						</xsl:for-each>
						<!-- 主要選單 End-->
						</div>
					</div>
					</td>
		<!-- 左方功能選單區 End-->

						<!-- 中央主資料區 Start-->
					<td id="center">

					<!-- accesskey start -->
					<div class="accesskey"><a href="accesskey.htm" title="中央區塊" accesskey="C">:::</a></div>
					<!-- accesskey start -->	

					<!-- friendly Start -->

					<div class="Path">
						<p><xsl:apply-templates select="//xPath" /></p>
					</div>
					
					<h4><xsl:value-of select="//xPath/UnitName" /></h4>
					<!--title end-->
					
					<xsl:apply-templates select="CatList" />
					<!--資料大類 END-->
					
					<xsl:apply-templates select="." mode="pageissue" />
					<!--分頁 END-->
				<table class="inline" summary="排版表格">
					<xsl:apply-templates select="TopicList" />
				</table>
					<!--標題列 END-->
					
					
					<!--分頁 END-->
					<div class="top"><a><xsl:attribute name="href">dp.asp?mp=<xsl:value-of select="mp" /></xsl:attribute>首頁</a></div>
					
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
			<xsl:if test="Article[@addtr='0']">
			<tr>
				<xsl:for-each select="Article[@addtr='0']">					
					
					<xsl:apply-templates select="." />					
												
				</xsl:for-each>
			</tr>
			</xsl:if>
			
			<xsl:if test="Article[@addtr='1']">
			<tr>
				<xsl:for-each select="Article[@addtr='1']">					
					
					<xsl:apply-templates select="." />					
												
				</xsl:for-each>
			</tr>
			</xsl:if>
			
			<xsl:if test="Article[@addtr='2']">
			<tr>
				<xsl:for-each select="Article[@addtr='2']">					
					
					<xsl:apply-templates select="." />					
												
				</xsl:for-each>
			</tr>
			</xsl:if>
			
			<xsl:if test="Article[@addtr='3']">
			<tr>
				<xsl:for-each select="Article[@addtr='3']">					
					
					<xsl:apply-templates select="." />					
												
				</xsl:for-each>
			</tr>
			</xsl:if>
			
			<xsl:if test="Article[@addtr='4']">
			<tr>
				<xsl:for-each select="Article[@addtr='4']">					
					
					<xsl:apply-templates select="." />					
												
				</xsl:for-each>
			</tr>
			</xsl:if>			
				
	</xsl:template>

	<xsl:template match="TopicList/Article">
		
			<td>
			<!--PIC START-->
				<xsl:if test="@IsPic ='Y'">							
					<xsl:if test="xImgFile">
						<img class="listimg">
							<xsl:attribute name="src">
								<xsl:value-of select="xImgFile" />
							</xsl:attribute>
							<xsl:attribute name="alt">
								<xsl:value-of select="Caption" />
							</xsl:attribute>
						</img>
					</xsl:if>
				</xsl:if>
			<!--PIC END-->
				<!--TITLE start-->
				<p><xsl:if test="xURL !=''">
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
			<!--TITLE END-->
			<br />
			<!--PostDate start-->			
			<xsl:if test="@IsPostDate='Y'">
				<xsl:value-of select="PostDate" />
			</xsl:if>
			<!--postdate end-->
			</p>			
			</td>			
			
	</xsl:template>
<!-- 主資料區-->

<!-- 分頁-->
	<xsl:template match="hpMain" mode="pageissue">
		<div class="term"> 共 <xsl:value-of select="totRec" /> 筆資料，第 <xsl:value-of select="nowPage" />/<xsl:value-of select="totPage" /> 頁，
		
	<xsl:if test="nowPage > 1">
	 <a class="previous"><xsl:attribute name="href">lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" /></xsl:attribute>
	 上一頁 </a> 
	 </xsl:if>
			 到第
	 		 <select id="pickPage" onChange="pageChange()">
	<script>
	var totPage = <xsl:value-of select="totPage" />;
	var qURL = "<xsl:value-of select="qURL" />";
	var perPage = <xsl:value-of select="PerPageSize" />
<![CDATA[
	for (xi=1;xi<=totPage;xi++)
		document.write("<option value=" +xi + ">" + xi + "</option>");
	

	function pageChange() {
		goPage(pickPage.value);
	}
	function goPage(nPage) {
		document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + perPage 
	}
	function perPageChange() {
		document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage 
	}
]]>	  
</script>
			</select>
	 頁 
	 
	 <xsl:if test="number(nowPage) &lt; number(totPage)">
	 	<a class="next"><xsl:attribute name="href">lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" /></xsl:attribute>
	 下一頁 </a> 
	 </xsl:if>	 
			<!--xsl:if test="queryYN"-->
			<div class="search"><A><xsl:attribute name="href">qp.asp?<xsl:value-of select="xqURL" /></xsl:attribute>
				條件查詢
			</A></div>
			<!--/xsl:if-->
<script>
	pickPage.value = <xsl:value-of select="nowPage" />;
	
</script>
	</div> <!--class="Page"-->
	</xsl:template>
<!-- 分頁-->


</xsl:stylesheet>
