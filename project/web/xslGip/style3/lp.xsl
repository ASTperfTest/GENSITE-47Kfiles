<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
<xsl:include href="info.xsl"/>

<xsl:template match="hpMain">
	<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
		<head>
			<xsl:apply-templates select="." mode="Header"/>
		</head>
		<script language="javascript" type="text/javascript">
			function bodyOnload() {
			    loadtagcloud();
			}
			function loadtagcloud() {
			var url="/services/tagCloud.aspx?cmd=gettagcloud";
			var myAjax = new Ajax(url, {
					method: 'get',
					update: $('tagcloud')
			}).request(); 		
         }
		</script>
		<link rel="stylesheet" type="text/css">
			<xsl:attribute name="href">../xslgip/<xsl:value-of select="myStyle"/>/css<xsl:value-of select="//MpStyle/Article/Caption"/>/member.css</xsl:attribute>
		</link>
		<script type="text/javascript" src="../mootools.v1.11.js"></script>
		<body onload="bodyOnload();">
			<!--header star (style A)-->
			<xsl:apply-templates select="." mode="RunActive"/>
			<!--header End-->
			<!--nav star (style A)-->
			<xsl:apply-templates select="." mode="toplink"/>
			<!--nav end-->
			<!--subnav star (style G)-->
			<xsl:apply-templates select="." mode="weblink"/>
			<!--subnav end-->
		
			<!--layout Table star (style A)-->
			<table class="layout">
			  <tr>
				<td class="left">
					<xsl:apply-templates select="." mode="l"/>
					<!--xsl:apply-templates select="." mode="TreeScriptCode"/-->
					<!--season star-->
					<xsl:apply-templates select="." mode="SolarTerms"/>
					<!--season end-->
					<!--menu star (style C)-->
					<xsl:apply-templates select="." mode="Menu"/>
					<!--menu end-->
					<!--RSS star-->
					<xsl:apply-templates select="." mode="xsRSS"/>
					<!--RSS end-->
					<!--ad star style A-->
					<xsl:apply-templates select="." mode="xsAD"/>
					<!--ad end-->
				</td>
				<td class="center">
					<xsl:apply-templates select="." mode="c"/>
					<!--path star (style B)-->
					<xsl:apply-templates select="." mode="xsPath"/>
					<!--path end-->
				<h3><xsl:value-of select="//xPath/UnitName"/></h3>
				
				<!-- 最新發問TAB LIST／開始 -->
						<!--分頁/開始 -->
						<xsl:apply-templates select="." mode="pageissue"/>		
						<!--分頁/結束 -->
						<!--資料大類/開始 -->

						<xsl:apply-templates select="CatList" />
						

						<!--List Start (A)-->
						<xsl:apply-templates select="TopicList" />
						<!--List End-->	    
				<!-- 最新發問TAB LIST／結束 -->
				<!--hotessay End-->    
				</td>
				<td class="right">
					<xsl:apply-templates select="." mode="r"/>
					<!--search star (style C)-->
					<xsl:apply-templates select="." mode="xsSearch"/>
					<!--search end-->
					<!--login star-->
					<xsl:apply-templates select="." mode="xsMember"/>
					<!--login end-->
					<!--epaper star (style B)-->
					<xsl:apply-templates select="." mode="xsePaper"/>
					<!--epaper end-->
					<!--tag Cloud -->
					<div class="TagCloud">
					    <h2>熱門關鍵字</h2>
					    <div id ="tagcloud">
						</div>
					</div>
					<!--ag Cloud end-->
				</td>
			  </tr>
			</table>
			<!--layout Table End-->
			<!--footer star (style A)-->
			<div class="footer">
				<xsl:apply-templates select="." mode="xsFooter"/>		
				<xsl:apply-templates select="." mode="xsCopyright"/>		
			</div>
			<!--footer End-->
		</body>
</html>
</xsl:template>


<!-- 主資料區-->
<xsl:template match="TopicList">
	<div class="Preface">
	<xsl:value-of disable-output-escaping="yes" select="../HeaderPart" />
	</div>
	<div class="list">
		<ul>
			<xsl:for-each select="Article">
				<li>
					<xsl:if test="xURL !=''">
						<a>
							<xsl:attribute name="href"><xsl:value-of select="xURL" /></xsl:attribute>
							<xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute>
							<xsl:if test="@newWindow='Y'"><xsl:attribute name="target">_blank</xsl:attribute></xsl:if>
							<xsl:value-of select="Caption" />
						</a>
					</xsl:if>
					<xsl:if test="xURL =''"><xsl:value-of select="Caption" /></xsl:if>
					<span class="from"><xsl:value-of select="DeptName" /></span>
					<span class="date"><xsl:value-of select="PostDate" /></span>
				</li>			
			</xsl:for-each>
		</ul>
	</div>
	<div class="Afterword">
	<xsl:value-of disable-output-escaping="yes" select="../FooterPart" />
	</div>
</xsl:template>
<!-- 主資料區-->

<!-- 資料大類 -->
<xsl:template match="CatList">
<div class="category">
<ul>
<li><a><xsl:attribute name="href">lp.asp?<xsl:value-of select="../xqURL" /></xsl:attribute>All</a></li>
    		<xsl:for-each select="CatItem">
    			<li>
    			<a>
    				<xsl:attribute name="href">lp.asp?<xsl:value-of select="../../xqURL" /><xsl:value-of select="xqCondition" /></xsl:attribute>
    				<xsl:value-of select="CatName" />
    			</a>
    			</li>
    		</xsl:for-each>
</ul>
</div>
</xsl:template>
<!-- 資料大類 -->
<!-- 資料次類 -->
<xsl:template match="CatListRelated">
</xsl:template>
<!-- 資料次類 -->
<!-- 分頁-->
<xsl:template match="hpMain" mode="pageissue">
	<xsl:variable name="PerPageSize"><xsl:value-of select="PerPageSize" /></xsl:variable>
	<!--xsl:if test="totRec>$PerPageSize"-->
		<div class="Page"> 
			第 <span class="Number"><xsl:value-of select="nowPage" />/<xsl:value-of select="totPage" /></span> 頁，共 <span class="Number"><xsl:value-of select="totRec" /></span> 筆
			<xsl:if test="nowPage > 1">
				<a title="上一頁">
					<xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute><img alt="上一頁"><xsl:attribute name="src">xslgip/<xsl:value-of select="//myStyle"/>/images/arrow_left.gif</xsl:attribute></img>上一頁
				</a>
			</xsl:if>			
			 &amp;nbsp;，到第
			<select name="pickPage" id="pickPage" class="inputtext" onChange="pageChange(this.value)">
				<script>
					var totPage = <xsl:value-of select="totPage" />;
					var qURL = "<xsl:value-of select="qURL" />";
					var nowPage = "<xsl:value-of select="nowPage" />";
					var PerPageSize = "<xsl:value-of select="PerPageSize" />";
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

			<xsl:if test="number(nowPage) &lt; number(totPage)">
				<a title="下一頁">
					<xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute><img alt="下一頁"><xsl:attribute name="src">xslgip/<xsl:value-of select="//myStyle"/>/images/arrow_right.gif</xsl:attribute></img>下一頁
				</a>
			</xsl:if>
			&amp;nbsp; ，每頁
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
				<option selected="selected">
					<xsl:attribute name="value">0<xsl:value-of select="PerPageSize" /></xsl:attribute>
					請選擇
				</option>
				<option value="15">15</option>
				<option value="30">30</option>
				<option value="50">50</option>
			</select> <label for="select">筆</label>資料
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
			<noscript>
				<br/>每頁<xsl:value-of select="PerPageSize" />筆,目前在第<xsl:value-of select="nowPage" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br/><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
	<!--/xsl:if-->
</xsl:template>
<!-- 分頁-->
</xsl:stylesheet>
