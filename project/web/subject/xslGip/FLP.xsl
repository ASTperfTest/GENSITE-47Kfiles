<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:user="urn:user-namespace-here" version="1.0">

  <xsl:template match="TopicList">
		<xsl:value-of disable-output-escaping="yes" select="../HeaderPart" />
		<!--標准文章樣式-->
		<TABLE border="0" align="center" cellpadding="3" cellspacing="1" class="body" xmlns="">
		 <TBODY>
		<xsl:apply-templates select="../TopicTitleList" />		
    	    	<xsl:for-each select="Article">
    	            <xsl:apply-templates select="."/>
    	    	</xsl:for-each>
    	    	</TBODY>
  		</TABLE>
		<xsl:value-of disable-output-escaping="yes" select="../FooterPart" />
    </xsl:template>
    
  <xsl:template match="TopicTitleList">    
   
  	<TR bgcolor="#BDDAE8">
	<xsl:for-each select="TopicTitle"> 			
  		<TH>
  		<xsl:value-of select="." />
		</TH>
    	</xsl:for-each>			
	</TR>

  </xsl:template>
   
  <xsl:template match="TopicList/Article">    
  	<TR bgcolor="#E7E7E7">
    	    	<xsl:for-each select="ArticleField">
    	            <xsl:apply-templates select="."/>
    	    	</xsl:for-each>
	</TR>
  </xsl:template>

  <xsl:template match="ArticleField">    
	<xsl:choose>	
		<xsl:when test="xURL">
			<TD>
    	        		<a>
		  		<xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
		  		<xsl:attribute name="title"><xsl:value-of select="Value" /></xsl:attribute> 
				<xsl:if test="xURL[@newWindow='Y']">
					<xsl:attribute name="target">_nwGip</xsl:attribute>
				</xsl:if>
		  		<xsl:value-of select="Value" />       	        		
    	        		</a>			
			</TD>
		</xsl:when>
		<xsl:otherwise>
			<TD>
  				<xsl:value-of select="Value" />    	        				
			</TD>
		</xsl:otherwise>
	</xsl:choose>	
  </xsl:template>

     <xsl:template match="hpMain" mode="pageissue">
	<div id="pageissue" class="Page">第 <xsl:value-of select="nowPage" />/<xsl:value-of select="totPage" /> 頁| 共 <xsl:value-of select="totRec" /> 筆| 跳至第
	  <select id="pickPage" class="inputtext" onChange="pageChange()">

<script>
	var totPage = <xsl:value-of select="totPage" />;
	var qURL = "<xsl:value-of select="qURL" />";
	
<![CDATA[
	for (xi=1;xi<=totPage;xi++)
		document.write("<option value=" +xi + ">" + xi + "</option>");
	

	function pageChange() {
		goPage(pickPage.value);
	}
	function goPage(nPage) {
		document.location.href= "flp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + perPage.value 
	}
	function perPageChange() {
		document.location.href= "flp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value 
	}
]]>	  
</script>
	  </select>
	 頁 
	 <xsl:if test="nowPage > 1">
	 | <a><xsl:attribute name="href">flp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" /></xsl:attribute>
	 上一頁 </a> 
	 </xsl:if>
	 <xsl:if test="number(nowPage) &lt; number(totPage)">
	 | <a><xsl:attribute name="href">flp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" /></xsl:attribute>
	 下一頁 </a> 
	 </xsl:if>
	 | 每頁筆數: 
	 <select name="perPage" class="inputtext" onChange="perPageChange()">
	   <option value="15">15</option>
	   <option value="20">20</option>
	   <option value="30">30</option>
	 </select>
			&amp;nbsp;　
			<xsl:if test="queryYN">
			<A><xsl:attribute name="href">qp.asp?<xsl:value-of select="xqURL" /></xsl:attribute>
				條件查詢
			</A>
			</xsl:if>
<script>
	pickPage.value = <xsl:value-of select="nowPage" />;
	perPage.value = <xsl:value-of select="PerPageSize" />;
</script>
	</div><!--id="pageissue"-->
     </xsl:template>  

</xsl:stylesheet>

