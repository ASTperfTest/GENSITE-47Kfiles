<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:saxon="http://icl.com/saxon/"
    extension-element-prefixes="saxon" version="1.0">
    <xsl:include href="info.xsl" />
    <xsl:template match="hpMain">
        <html xmlns="http://www.w3.org/1999/xhtml">
            <head>
                <xsl:apply-templates select="." mode="topinfo" />
                <link rel="stylesheet" type="text/css">
                    <xsl:attribute name="href">/xslGip/<xsl:apply-templates select="//MpStyle" />/css/mp.css</xsl:attribute>
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
                            <!-- 廣告區／開始 -->
                            <xsl:apply-templates select="AdRotate" />
                            <!-- 廣告區／結束 -->
                        </td>
                        <!-- 左欄／結束 -->
                        <!-- 中欄／開始 -->
                        <td class="Center">
                            <xsl:apply-templates select="." mode="c" />
                            <!-- 認識蝴蝶蘭與蝴蝶蘭小故事／開始 -->
                            <table border="0" cellspacing="0" cellpadding="0" class="News0" summary="排版表格">
                                <tr>
                                    <td class="left">
                                        <xsl:apply-templates select="BlockA" />
                                    </td>
                                    <td class="right">
                                        <xsl:apply-templates select="BlockB" />
                                    </td>
                                </tr>
                            </table>
                            <!-- 認識蝴蝶蘭與蝴蝶蘭小故事／結束 -->
                            <!-- 館務資訊／開始 -->
                            <xsl:apply-templates select="BlockC" />
			    
                            <!-- 館務資訊／結束 -->
                            <!-- 攝影藝廊／開始 -->
                            <xsl:apply-templates select="BlockD" />
                            <!-- 攝影藝廊／結束 -->
                            <!-- 蝴蝶蘭與我／開始 -->
                            <xsl:apply-templates select="BlockE" />
                            <!-- 蝴蝶蘭與我／結束 -->
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
    <!-- 認識蝴蝶蘭與蝴蝶蘭小故事／開始 -->
    <xsl:template match="BlockA">
        <h1>
            <xsl:value-of select="Caption" />
        </h1>
        <xsl:for-each select="Article">
            <p>
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
	      		<xsl:value-of disable-output-escaping="yes" select="substring(Content, 1, 120)" />...
                    <!--xsl:value-of disable-output-escaping="yes" select="Content" /-->
                </a>
            </p>
        </xsl:for-each>
    </xsl:template>
    <xsl:template match="BlockB">
        <h1>
            <xsl:value-of select="Caption" />
        </h1>
        <xsl:for-each select="Article">
            <p>
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
			<xsl:value-of disable-output-escaping="yes" select="substring(Content, 1, 120)" />...
                    <!--xsl:value-of disable-output-escaping="yes" select="Content" /-->
                </a>
            </p>
        </xsl:for-each>
    </xsl:template>
    <!-- 認識蝴蝶蘭與蝴蝶蘭小故事／結束 -->
    <!-- 館務資訊／開始 -->
    <xsl:template match="BlockC">
        <div class="News1">
            <h3>
                <xsl:value-of select="Caption" />
            </h3>
            <div class="Body">
                <!--xsl:for-each select="Article[position()&lt;=2]"-->
		<xsl:for-each select="Article">
                    <h4>
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
                    </h4>
                    <xsl:if test="xImgFile">
                        <a target="_nwGip">
                            <xsl:attribute name="href">
                                <xsl:value-of select="//MpPublicPath" />/<xsl:value-of select="xImgFile" />
                            </xsl:attribute>
                            <xsl:attribute name="title">
                                <xsl:value-of select="Caption" />
                            </xsl:attribute>
                            <img border="0">
                                <xsl:attribute name="src">
                                    <xsl:value-of select="//MpPublicPath" />/<xsl:value-of select="xImgFile" />
                                </xsl:attribute>
                                <xsl:attribute name="alt">
                                    <xsl:value-of select="Caption" />
                                </xsl:attribute>
                            </img>
                        </a>
                    </xsl:if>
                    <p>
                        <xsl:value-of disable-output-escaping="yes" select="Content" />
                    </p>
                </xsl:for-each>
                <ul>
                    <!--xsl:for-each select="//BlockC2/Article[position()&gt;2]"-->
		    <xsl:for-each select="//BlockC2/Article">
                        <li>
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
                            <span class="Date">
                                <xsl:value-of select="PostDate" />
                            </span>
                        </li>
                    </xsl:for-each>
                </ul>
                <div class="More">
                    <a>
					    <xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="//BlockC2/@xNode" /></xsl:attribute> 
					    <xsl:attribute name="title">更多內容</xsl:attribute>
					    更多內容
				    </a>
                </div>
            </div>
            <div class="Foot"></div>
        </div>
    </xsl:template>
    <!-- 館務資訊／結束 -->
    <!-- 活動照片／開始 -->
    <xsl:template match="BlockD">
        <div class="News2">
            <div class="More1">
                <a>
                    <xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="@xNode" /></xsl:attribute>
                    <xsl:attribute name="title">更多內容</xsl:attribute>
                    <img src="images/icon_more01.gif" border="0">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/icon_more01.gif</xsl:attribute>
                    </img>
                </a>
            </div>
            <h3>
                <xsl:value-of select="Caption" />
            </h3>
            <div class="Body">
                <xsl:for-each select="Article/xImgFile">
                    <a target="_nwGip">
                        <xsl:attribute name="href">
                            <xsl:value-of select="//MpPublicPath" />/<xsl:value-of select="." />
                        </xsl:attribute>
                        <xsl:attribute name="title">
                            <xsl:value-of select="../Caption" />
                        </xsl:attribute>
                        <img border="0" class="photo">
                            <xsl:attribute name="src">
                                <xsl:value-of select="//MpPublicPath" />/<xsl:value-of select="." />
                            </xsl:attribute>
                            <xsl:attribute name="alt">
                                <xsl:value-of select="../Caption" />
                            </xsl:attribute>
                        </img>
                    </a>
                    <xsl:if test="position() mod 4 = 0">
                        <br />
                    </xsl:if>
                </xsl:for-each>
            </div>
        </div>
    </xsl:template>
    <!-- 活動照片／結束 -->
    <!-- 蝴蝶蘭與我／開始 -->
    <xsl:template match="BlockE">
        <div class="News3">
            <div class="More1">
                <a>
                    <xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="@xNode" /></xsl:attribute>
                    <xsl:attribute name="title">更多內容</xsl:attribute>
                    <img src="images/icon_more01.gif" border="0">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/icon_more01.gif</xsl:attribute>
                    </img>
                </a>
            </div>
            <h3>
                <xsl:value-of select="Caption" />
            </h3>
            <div class="Body">
                <table summary="排版表格">
                    <tr>
                        <xsl:for-each select="Article">
                            <td valign="top">
                                <xsl:choose>
                                    <xsl:when test="position() mod 2 = 1">
                                        <xsl:attribute name="class">left</xsl:attribute>
                                    </xsl:when>
                                    <xsl:otherwise>
                                        <xsl:attribute name="class">right</xsl:attribute>
                                    </xsl:otherwise>
                                </xsl:choose>
                                <xsl:if test="xImgFile">
                                    <a target="_nwGip">
                                        <xsl:attribute name="href">
                                            <xsl:value-of select="//MpPublicPath" />/<xsl:value-of select="xImgFile" />
                                        </xsl:attribute>
                                        <xsl:attribute name="title">
                                            <xsl:value-of select="Caption" />
                                        </xsl:attribute>
                                        <img border="0">
                                            <xsl:attribute name="src">
                                                <xsl:value-of select="//MpPublicPath" />/<xsl:value-of select="xImgFile" />
                                            </xsl:attribute>
                                            <xsl:attribute name="alt">
                                                <xsl:value-of select="Caption" />
                                            </xsl:attribute>
                                        </img>
                                    </a>
                                </xsl:if>
                                <p>
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
                                        <xsl:value-of select="Content" />
                                    </a>
                                </p>
                            </td>
                            <xsl:if test="position() mod 2 = 0">
                                <tr />
                            </xsl:if>
                        </xsl:for-each>
                    </tr>
                </table>
            </div>
            <div class="Foot"></div>
        </div>
    </xsl:template>
    <!-- 蝴蝶蘭與我／結束 -->
    <!-- 廣告區／開始 -->
    <xsl:template match="AdRotate">
        <div class="AD">
            <xsl:for-each select="Article">
                <a>
                    <xsl:attribute name="href">
                        <xsl:value-of select="xURL" />
                    </xsl:attribute>
                    <xsl:variable name="title">
                        <xsl:value-of select="Caption" />
                    </xsl:variable>
                    <xsl:if test="@newWindow='Y'">
                        <xsl:attribute name="target">_nwGip</xsl:attribute>
                    </xsl:if>
                    <xsl:if test="xImgFile">
                        <img border="0">
                            <xsl:attribute name="src">
                                <xsl:value-of select="//MpPublicPath" />/<xsl:value-of select="xImgFile" />
                            </xsl:attribute>
                            <xsl:attribute name="alt">
                                <xsl:value-of select="Caption" />
                            </xsl:attribute>
                        </img>
                    </xsl:if>
                </a>
            </xsl:for-each>
        </div>
    </xsl:template>
    <!-- 廣告區／結束 -->
</xsl:stylesheet>
