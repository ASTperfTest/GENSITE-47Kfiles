<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />

	<xsl:template match="BlockKW">	

		<xsl:value-of select="Caption" />：
		<ul>
			<xsl:for-each select="Article">
				<li>
					<a>
						<xsl:attribute name="href">#</xsl:attribute>
						<xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute>						
						<xsl:attribute name="onclick">javascript:doSearch('<xsl:value-of select="Caption" />')</xsl:attribute>						
						<xsl:value-of select="Caption" />
					</a>
				</li>
			</xsl:for-each>			
		</ul>

	<script language="JavaScript" type="text/JavaScript">
	<![CDATA[
	<!--
	function doSearch(value){
		document.getElementById("Keyword").value = value;
		document.getElementById("SearchForm").action = "/kp.asp?xdURL=Search/SearchResultList.asp&mp=1";
		document.getElementById("SearchForm").submit();
	}
	//-->
	]]>
	</script>
		<!--div class="more">
			<a>
				<xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="@xNode" />&amp;mp=<xsl:value-of select="//mp" /></xsl:attribute> 
				<xsl:attribute name="title">more</xsl:attribute>
				更多關鍵字
			</a>
		</div-->
	</xsl:template>

</xsl:stylesheet>
