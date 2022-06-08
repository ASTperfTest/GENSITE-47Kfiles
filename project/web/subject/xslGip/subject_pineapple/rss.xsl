<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:saxon="http://icl.com/saxon/"
    extension-element-prefixes="saxon" version="1.0">
    <xsl:include href="info.xsl" />
    <xsl:template match="hpMain">
        <html xmlns="http://www.w3.org/1999/xhtml">
            <head>
                <xsl:apply-templates select="." mode="topinfo" />
                <link rel="stylesheet" type="text/css">
                    <xsl:attribute name="href">/xslGip/<xsl:apply-templates select="//MpStyle" />/css/rss.css</xsl:attribute>
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
                                <xsl:value-of select="RSSTop/Caption" /> <!--xsl:value-of select="//CtUnitName"/--></div>
                            <!-- 單元名稱／結束 -->
                            <!-- 導覽列／開始 -->
                            <xsl:apply-templates select="." mode="RssxPath" />
                            <!-- 導覽列／結束 -->
                            <!-- Rss 網站導覽／開始 -->
                            <div class="RSS">
                                <xsl:apply-templates select="." mode="RSSTop" />
                                <xsl:apply-templates select="." mode="RssSitemap" />
                            </div>
                            <!-- Rss 網站導覽／結束 -->
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
</xsl:stylesheet>
