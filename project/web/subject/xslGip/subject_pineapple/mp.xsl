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
                            <xsl:apply-templates select="." mode="Menu2" />
                            <!-- 選單／結束 -->
                            <!-- 廣告區／開始 -->
                            <xsl:apply-templates select="AdRotate" />
                            <!-- 廣告區／結束 -->
                        </td>
                        <!-- 左欄／結束 -->
                        <!-- 中欄／開始 -->
                        <td class="Center">
                            <xsl:apply-templates select="." mode="c" />
                            <!-- 認識鳳梨／開始 -->
                            <div id="News0">
                            <table border="0" cellspacing="0" cellpadding="0" class="layout01" summary="排版表格">
                                <tr>
                                    <td class="left">
                                        <xsl:apply-templates select="BlockA" />
                                    </td>
                                    <td class="right"></td>
                                                                    </tr>
                            </table>
                            </div>
                            <!-- 認識鳳梨／結束 -->
 
                           <!-- 鳳梨達人／開始 -->
                            <xsl:apply-templates select="BlockB" />
                            <!-- 鳳梨達人／結束 -->

                            <div class="News2">
                            <table class="layout2" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                            <td scope="col" class="left">
                            <!-- 攝影藝廊／開始 -->
                            <xsl:apply-templates select="BlockPhoto" />
                            <!-- 攝影藝廊／結束 -->
                            </td>
                            <td scope="col" class="right">
                            <!-- 你不知道的鳳梨／開始 -->
                            <xsl:apply-templates select="BlockE" />
                            <!-- 你不知道的鳳梨／結束 -->
                            </td>
                            </tr>
                            </table>
                            </div>
 
                            <!-- 鳳梨消息／開始 -->
                            <xsl:apply-templates select="BlockNews" />
                            <!-- 鳳梨消息／結束 -->
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
    <!-- 認識鳳梨／開始 -->
    <xsl:template match="BlockA">
           <h1>
               <xsl:value-of select="Caption" />
           </h1>
           <xsl:for-each select="Article">
                   <xsl:if test="xImgFile">
                       <a target="_nwGip">
                           <xsl:attribute name="href">
                               <xsl:value-of select="//MpPublicPath" />/<xsl:value-of select="xImgFile" />
                           </xsl:attribute>
                           <xsl:attribute name="title">
                               <xsl:value-of select="Caption" />
                           </xsl:attribute>
                           <img border="0" class="img" >
                               <xsl:attribute name="src">
                                   <xsl:value-of select="//MpPublicPath" />/<xsl:value-of select="xImgFile" />
                               </xsl:attribute>
                               <xsl:attribute name="alt">
                                   <xsl:value-of select="Caption" />
                               </xsl:attribute>
                           </img>
                       </a>
                   </xsl:if>
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
                   <p>
                       <xsl:value-of disable-output-escaping="yes" select="Content" />
                   </p>
               </xsl:for-each>
    </xsl:template>
    <!-- 認識鳳梨／結束 -->

    <!-- 鳳梨達人／開始 -->
    <xsl:template match="BlockB">
        <div class="News1">
            <h3>
                <xsl:value-of select="Caption" />
            </h3>
            <div class="Body">
		<xsl:for-each select="Article">
		<xsl:if test="position() &lt; 3">
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
                    <p>
                        <xsl:value-of disable-output-escaping="yes" select="Abstract" />
                    </p>
		<div style="clear:both"></div>
		  </xsl:if>
                </xsl:for-each>
                <div style="clear:both"></div>
                <ul>
        
		    <xsl:for-each select="//BlockB/Article">
		    <xsl:if test="position() > 2">
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
          </xsl:if>
                    </xsl:for-each>
                </ul>
                <div class="More">
                <a>
		<xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="//BlockB/@xNode" /></xsl:attribute> 
		<xsl:attribute name="title">更多內容</xsl:attribute>
		更多內容
		</a>
                </div>
            </div>
            <div class="Foot"></div>
        </div>
    </xsl:template>
    <!-- 鳳梨達人／結束 -->

    <!-- 攝影藝廊／開始 -->
    <xsl:template match="BlockPhoto">
            <div class="More2">
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
                <xsl:for-each select="Article">
		<div class="floatb">
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
                            <img border="0" class="img">
                                <xsl:attribute name="src">
                                    <xsl:value-of select="//MpPublicPath" />/<xsl:value-of select="xImgFile" />
                                </xsl:attribute>
                                <xsl:attribute name="alt">
                                    <xsl:value-of select="Caption" />
                                </xsl:attribute>
                            </img>
                        </a>
                    </xsl:if>
		</div>
                </xsl:for-each>
    </xsl:template>
    <!-- 攝影藝廊／結束 -->

    <!-- 你不知道的鳳梨／開始 -->
    <xsl:template match="BlockE">
            <h3>
                <xsl:value-of select="Caption" />
            </h3>
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
                   <p>
                       <xsl:value-of disable-output-escaping="yes" select="Content" />
                   </p>
           </xsl:for-each>
    </xsl:template>
    <!-- 你不知道的鳳梨／結束 -->

    <!-- 鳳梨消息／開始 -->
    <xsl:template match="BlockNews">
        <div class="News3">
            <div class="More1">
                <a>
		<xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="//BlockNews/@xNode" /></xsl:attribute> 
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
                <ul>
		    <xsl:for-each select="//BlockNews/Article">
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
            </div>
            <div class="Foot"></div>
        </div>
    </xsl:template>
    <!-- 鳳梨消息／結束 -->

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
