<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:saxon="http://icl.com/saxon/"
    extension-element-prefixes="saxon" version="1.0">
    <xsl:include href="info.xsl" />
    <xsl:include href="pageissue.xsl" />
    <xsl:template match="hpMain">
        <html xmlns="http://www.w3.org/1999/xhtml">
            <head>
                <xsl:apply-templates select="." mode="topinfo" />
                <link rel="stylesheet" type="text/css">
                    <xsl:attribute name="href">/xslGip/<xsl:apply-templates select="//MpStyle" />/css/lp.css</xsl:attribute>
                </link>
            </head>
            <body>
                <!-- 網站標題／開始 -->
                <xsl:apply-templates select="." mode="webinfo" />
                <!-- 網站標題／結束 -->
                <!-- 導覽列／開始 -->
                <xsl:apply-templates select="." mode="info" />
                <!-- 導覽列／結束 -->
                <!-- 首頁排版表格／開始 -->
                <table class="Layout" summary="排版表格">
                    <tr>
                        <!-- 左欄／開始 -->
                        <td class="Left">
                            <xsl:apply-templates select="." mode="m" />
                            <!-- 選單／開始 -->
                            <xsl:apply-templates select="." mode="Menu" />
                            <!-- 選單／結束 -->
                        </td>
                        <!-- 左欄／結束 -->
                        <!-- 中欄／開始 -->
                        <td class="Center">
                            <xsl:apply-templates select="." mode="c" />
                            <!-- 單元名稱／開始 -->
                            <div class="PageTitle">
                                <xsl:value-of select="//xPath/UnitName" />
                            </div>
                            <!-- 單元名稱／結束 -->
                            <!-- RSS，查詢本單元／開始 -->
                            <xsl:apply-templates select="." mode="xsFunctions" />
                            <!-- RSS，查詢本單元／結束 -->
                            <!-- 瀏覽路徑／開始 -->
                            <xsl:apply-templates select="." mode="xsPath" />
                            <!-- 瀏覽路徑／結束 -->
                            <!-- 分頁／開始 -->
                            <xsl:apply-templates select="." mode="pageissueC" />
                            <!-- 分頁／結束 -->
			    <!-- 投稿說明／開始 -->
			    <div class="Page"><br/>
			  	  <xsl:if test="//xPath/UnitName='蝴蝶蘭與我' or //xPath/UnitName='常見問答'">
					<img src="xslgip/subject_Phalaenopsis/images/mailus.gif" alt="我要投稿" align="absmiddle" />
					<a href="ct.asp?xItem=6666&amp;ctNode=1415&amp;mp=5" title="我要投稿">我要投稿</a>
			          </xsl:if>
			    </div>
			    <!-- 投稿說明／結束 -->
                            <!-- 資料大類／開始 -->
                            <xsl:apply-templates select="." mode="xsCatList" />
                            <!-- 資料大類／結束 -->
                            <xsl:apply-templates select="TopicList" />
			    
                        </td>
                        <!-- 中欄／結束 -->
                    </tr>
                </table>
                <!-- 首頁排版表格／結束 -->
                <!-- 頁尾資訊／開始 -->
                <div class="Bottom">
                    <xsl:apply-templates select="." mode="bottominfo" />
                </div>
                <!-- 頁尾資訊／結束 -->
            </body>
        </html>
    </xsl:template>
    <!-- 主資料區／開始-->
    <xsl:template match="TopicList">
        <xsl:value-of disable-output-escaping="yes" select="../HeaderPart" />
        <ol class="oList">
            <xsl:for-each select="Article">
                <li>
                    <xsl:if test="xURL !=''">
                        <span class="Title">
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
                        </span>
                    </xsl:if>
                    <xsl:if test="xURL =''">
                        <span class="Title">
                            <xsl:value-of select="Caption" />
                        </span>
                    </xsl:if>
                    <span class="Date">
                        <xsl:value-of select="PostDate" />
                    </span>
                    <span class="Brief">
                        <xsl:value-of select="Abstract" />
                    </span>
                </li>
            </xsl:for-each>
        </ol>
        <xsl:value-of disable-output-escaping="yes" select="../FooterPart" />
    </xsl:template>
    <!-- 主資料區／結束-->
    <!-- 分頁／開始-->
    <xsl:template match="hpMain" mode="pageissue">
        <xsl:variable name="PerPageSize">
            <xsl:value-of select="PerPageSize" />
        </xsl:variable>
        <!--xsl:if test="totRec>$PerPageSize"-->
        <div class="Page"> 
			共 <span class="Number">
                <xsl:value-of select="totPage" />
            </span> 頁， <span class="Number">
                <xsl:value-of select="totRec" />
            </span> 筆資料
			<xsl:if test="nowPage > 1">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="上一頁">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_prev.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>			
			 &amp;nbsp;到第
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
			頁

			<xsl:if test="number(nowPage) &lt; number(totPage)">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="下一頁">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_next.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>
			&amp;nbsp; 每頁顯示
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
                <option selected="selected">
					<xsl:attribute name="value">0<xsl:value-of select="PerPageSize" /></xsl:attribute>
					請選擇
				</option>
                <option value="15">15</option>
                <option value="30">30</option>
                <option value="50">50</option>
            </select> 筆
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
			<noscript>
				<br />每頁<xsl:value-of select="PerPageSize" />筆,目前在第<xsl:value-of select="nowPage" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br /><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
        <!--/xsl:if-->
    </xsl:template>
    <!-- 分頁／結束-->
</xsl:stylesheet>
