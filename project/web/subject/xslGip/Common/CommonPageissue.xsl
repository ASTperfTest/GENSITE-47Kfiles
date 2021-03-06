<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<!-- 分頁-->
	<xsl:template match="hpMain" mode="pageissue">
		<script>
			var totPage = <xsl:value-of select="totPage" />;
			var qURL = "<xsl:value-of select="qURL" />";
			totPage = totPage == 0 ? 1 : totPage;
		</script>
		<div class="term">
			共 <xsl:value-of select="totRec" /> 筆資料，第 <xsl:value-of select="nowPage" />/<script>document.write(totPage)</script> 頁，
			每頁顯示
			<select name="perPage" onChange="perPageChange()">
				<option value="20">20</option>
				<option value="50">50</option>
				<option value="100">100</option>
			</select>筆，
			<xsl:if test="nowPage > 1">
				<a class="previous">
					<xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
					上一頁
				</a>
			</xsl:if>
			到第
			<select id="pickPage" onChange="pageChange()">
				<script>
<![CDATA[
	for (xi=1;xi<=totPage;xi++)
		document.write("<option value=" +xi + ">" + xi + "</option>");
	

	function pageChange() {
		goPage(pickPage.value);
	}
	function goPage(nPage) {
		document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + perPage.value 
	}
	function perPageChange() {
		document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value 
	}
]]>
				</script>
			</select>
			頁

			<xsl:if test="number(nowPage) &lt; number(totPage)">
				<a class="next">
					<xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
					下一頁
				</a>
			</xsl:if>
			<!--xsl:if test="queryYN"-->
			<div class="search">
				<A>
					<xsl:attribute name="href">
						qp.asp?<xsl:value-of select="xqURL" />
					</xsl:attribute>
					條件查詢
				</A> 
				<!--若不鎖右鍵的話才顯示RSS-->
				<xsl:if test="//lockrightbtn='N'">| <img src="/subject/xslgip/rss.gif" alt="rss" /><a href="_blank">
					    <xsl:attribute name="href">
						rss.asp?xnode=<xsl:value-of select="//TopicList/@xNode" />
					    </xsl:attribute> RSS訂閱
				    </a>
				</xsl:if>
			</div>
			<!--/xsl:if-->
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
		</div>
		<!--class="Page"-->
	</xsl:template>
	<!-- 分頁-->
</xsl:stylesheet>






