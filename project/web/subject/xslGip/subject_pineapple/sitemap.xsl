<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
    xmlns:user="urn:user-namespace-here" version="1.0">
    <xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
    <xsl:include href="info.xsl" />
    <xsl:template match="hpMain">
        <html xmlns="http://www.w3.org/1999/xhtml">
            <head>
                <xsl:apply-templates select="." mode="topinfo" />
                
                <link rel="stylesheet" type="text/css">
                    <xsl:attribute name="href">/xslGip/<xsl:apply-templates select="//MpStyle" />/css/sitemap.css</xsl:attribute>
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
                                <xsl:value-of select="//xPath/UnitName" />網站導覽 
                            </div>
                            <!-- 單元名稱／結束 -->
                            <!-- 瀏覽路徑／開始 -->
                            <div class="Path" xmlns="">
            			瀏覽路徑：<a href="mp.asp">回首頁 </a>> 網站導覽
			    </div> 
                            <!-- 瀏覽路徑／結束 -->
                            
                            <div class="Sitemap">
				<h3>快速鍵設定說明</h3>
				<ul>
					<li>Alt+U：網站標題及導覽列。</li>
					<li>Alt+M：主選單。</li>
					<li>Alt+C：主要內容區。</li>
					<li>Alt+R：首頁右邊區塊，包括「訂閱電子報」、「活動照片」、「廣告區」等。</li>
					<li>Alt+S：站內查詢</li>
				</ul>

				<h3>網站導覽</h3>
				<xsl:apply-templates select="Sitemap" />
			    </div>
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


<xsl:template match="Sitemap">
        <ul>
            <xsl:for-each select="MenuCat">
                <li>
                    <a>
                        <xsl:attribute name="href">
                            <xsl:value-of select="xURL" />
                        </xsl:attribute>
                        <xsl:value-of select="Caption" />
                    </a>
                </li>
                <ul>
                    <xsl:for-each select="MenuCat">
                        <li>
                            <a>
                                <xsl:attribute name="href">
                                    <xsl:value-of select="xURL" />
                                </xsl:attribute>
                                <xsl:value-of select="Caption" />
                            </a>
                        </li>
                    </xsl:for-each>
                </ul>
            </xsl:for-each>
        </ul>
    </xsl:template>
    <!--Sitemap End-->

 


</xsl:stylesheet>
