<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
	<xsl:include href="headpart.xsl" />
	<xsl:include href="hySearch.xsl" />
	<xsl:include href="epaper.xsl"/>
	<xsl:include href="footer.xsl"/>
	<xsl:include href="menuitem.xsl"/>
	<xsl:include href="skeyword.xsl" />
	<xsl:include href="BlockA.xsl" />
	<xsl:include href="BlockB.xsl" />
	<xsl:include href="BlockC.xsl" />
	<xsl:include href="BlockD.xsl" />
	<xsl:include href="BlockE.xsl" />
	<xsl:include href="BlockF.xsl" />
	<xsl:include href="info.xsl" />
	
	<xsl:template match="hpMain">
		<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
			<xsl:apply-templates select="." mode="WInfo" />
		
			<body class="body">	
				<!-- 網頁標題 (headpart.xsl) Start-->
					<xsl:apply-templates select="." mode="headinfo" />
				<!-- 網頁標題 End-->	
					
				<!--中間區塊 start -->
					<div id="middle">
						<table class="layout" summary="排版表格">	
							<tr>
							<!-- leftcol start -->	
								<td class="leftbg">
								<div id="leftcol"> 
								<!-- accesskey start -->
								<div class="accesskey"><a href="accesskey.htm" title="左方區塊" accesskey="L">:::</a></div>
								<!-- accesskey start -->	

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
					<!-- leftcol end -->


<!-- content start -->
<td id="center">

<!-- accesskey start -->
<div class="accesskey"><a href="accesskey.htm" title="中央區塊" accesskey="C">:::</a></div>
<!-- accesskey start -->
<div class="skeyword">
					<xsl:apply-templates select="BlockKW" />
</div>
	
					<xsl:apply-templates select="BlockA" />
		<table class="column" summary="排版表格">
			<tr>
				<td class="left">
            <!-- BlockX Start-->				
					<xsl:apply-templates select="BlockB" />
				</td>
				<td class="right">
					<xsl:apply-templates select="BlockC" />
				</td>
			</tr>
		</table>				
					<xsl:apply-templates select="BlockD" />
					<xsl:apply-templates select="BlockE" />
					<xsl:if test="//BlockF/Article">
					<xsl:apply-templates select="BlockF" />
					</xsl:if>
			<!-- BlockX end -->

</td>
<!-- content end -->
</tr>
</table>
</div>	
<!--中間區塊 end -->
		
		<!-- 頁尾資訊 Start-->
				<xsl:apply-templates select="." mode = "Copyright"/>
				
		<!-- 頁尾資訊 End-->
			</body>
		</html>
	</xsl:template>
	


</xsl:stylesheet>
