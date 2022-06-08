<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:user="urn:user-namespace-here" version="1.0">

<xsl:template match="TopicList">
<xsl:variable name="pp"><xsl:value-of select="//PerPageSize_Game" /></xsl:variable>
<xsl:variable name="np"><xsl:value-of select="//nowPage_Game" /></xsl:variable>
	<div class="lp">
		<xsl:value-of disable-output-escaping="yes" select="../HeaderPart" />
		<!--table width="100%" border="0" cellspacing="0" cellpadding="0" summary="問題列表資料表格">
		<tr>
		<th>獎項</th><th>特獎 + 頭獎</th><th>壹獎 + 貳獎</th><th>參獎 + 普獎</th><th>參加獎</th>
		</tr>
		<tr>
		<td>目前中獎機率</td>
		 <xsl:for-each select="//GamePrize/PrizeRate">
		 	<td>
			<xsl:choose>
				<xsl:when test="Prize>100">
					100%
				</xsl:when>
				<xsl:otherwise>
					  <xsl:value-of select="Prize" />%
				</xsl:otherwise>
			</xsl:choose>
		 	  

      
      </td>
		 </xsl:for-each>
		</tr-->
		<!--tr>
		<td></td><td></td><td></td><td></td><td><a href="http://kmbeta.coa.gov.tw/iqcoa/gameprize.htm" target="_blank">活動獎項ㄧ覽</a></td>
		</tr>
		</table>
		<br/-->

		<br/>
    <table width="100%" border="0" cellspacing="0" cellpadding="0" summary="問題列表資料表格">
		<tr>
		<th>排名</th><th>姓名｜暱稱</th><th>得分</th>
		</tr>
			
				<xsl:for-each select="//coaGame/Rank">
					<tr>
						<td><xsl:value-of select="(number($np)-1)*number($pp)+position()"/></td>
						<td><xsl:value-of select="substring(Player,1,1)" />☆<xsl:value-of select="substring(Player,3)" />
						|<xsl:value-of select="Player2" /></td>
						<td><xsl:value-of select="Score" /></td>
					</tr>
				</xsl:for-each>
		</table>
		<xsl:value-of disable-output-escaping="yes" select="../FooterPart" />
	</div>
</xsl:template>
    
<xsl:template match="TopicTitleList">    
  	<tr>
		<th scope="col">&amp;nbsp;</th>
		<xsl:for-each select="TopicTitle"> 			
			<th scope="col"><xsl:value-of select="." /></th>
		</xsl:for-each>			
	</tr>
</xsl:template>

<xsl:template match="ArticleField">    
	<xsl:choose>	
		<xsl:when test="xURL">
			<td>
				<a>
					<xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="Value" /></xsl:attribute> 
					<xsl:if test="xURL[@newWindow='Y']">
						<xsl:attribute name="target">_blank</xsl:attribute>
					</xsl:if>
					<xsl:value-of select="Value" />       	        		
				</a>			
			</td>
		</xsl:when>
		<xsl:when test="Value=''"><td>&amp;nbsp;</td></xsl:when>
		<xsl:otherwise><td><xsl:value-of select="Value" /></td></xsl:otherwise>
	</xsl:choose>	
</xsl:template>

<!-- 分頁-->
<xsl:template match="hpMain" mode="pageissue">
	<xsl:variable name="PerPageSize"><xsl:value-of select="PerPageSize_Game" /></xsl:variable>
	<!--xsl:if test="totRec>$PerPageSize"-->
		<div class="Page"> 
			第 <span class="Number"><xsl:value-of select="nowPage_Game" />/<xsl:value-of select="totPage_Game" /></span> 頁，共 <span class="Number"><xsl:value-of select="totRec_Game" /></span> 筆
			<xsl:if test="nowPage_Game > 1">
				<a title="上一頁">
					<xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage_Game - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize_Game" />
					</xsl:attribute><img alt="上一頁"><xsl:attribute name="src">xslgip/<xsl:value-of select="//myStyle"/>/images/arrow_left.gif</xsl:attribute></img>上一頁
				</a>
			</xsl:if>			
			 &amp;nbsp;，到第
			<select name="pickPage" id="pickPage" class="inputtext" onChange="pageChange(this.value)">
				<script>
					var totPage = <xsl:value-of select="totPage_Game" />;
					var qURL = "<xsl:value-of select="qURL" />";
					var nowPage = "<xsl:value-of select="nowPage_Game" />";
					var PerPageSize = "<xsl:value-of select="PerPageSize_Game" />";
					<![CDATA[
					for (xi=1;xi<=totPage;xi++)
					if (xi==nowPage)
					document.write("<option value=" +xi + " selected='selected'>" + xi + "</option>");
					else
					document.write("<option value=" +xi + ">" + xi + "</option>");
					
					function pageChange(nPage) {
					//alert(nPage);
					//goPage(pickPage.value);
					goPage(nPage);
					}
					function goPage(nPage) {
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + PerPageSize
					}
					function perPageChange(pagesize) {
					//document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nowPage + "&pagesize=" + pagesize
					}
					]]>
				</script>
			</select>
			<label for="select">頁</label>，

			<xsl:if test="number(nowPage_Game) &lt; number(totPage_Game)">
				<a title="下一頁">
					<xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage_Game + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize_Game" />
					</xsl:attribute><img alt="下一頁"><xsl:attribute name="src">xslgip/<xsl:value-of select="//myStyle"/>/images/arrow_right.gif</xsl:attribute></img>下一頁
				</a>
			</xsl:if>
			&amp;nbsp; ，每頁
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
				<option selected="selected">
					<xsl:attribute name="value">0<xsl:value-of select="PerPageSize_Game" /></xsl:attribute>
					請選擇
				</option>
				<option value="15">15</option>
				<option value="30">30</option>
				<option value="50">50</option>
			</select> <label for="select">筆</label>資料
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage_Game" />;
				perPage.value = <xsl:value-of select="PerPageSize_Game" />;
			</script>
			<noscript>
				<br/>每頁<xsl:value-of select="PerPageSize_Game" />筆,目前在第<xsl:value-of select="nowPage_Game" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br/><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
	<!--/xsl:if-->
</xsl:template>
<!-- 分頁-->

</xsl:stylesheet>

