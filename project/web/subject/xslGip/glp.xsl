<?xml version="1.0" encoding="big5" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:saxon="http://icl.com/saxon/"
    extension-element-prefixes="saxon" version="1.0">
    <!-- 主資料區／開始-->
    <xsl:template match="TopicList">
        <link rel="stylesheet" type="text/css">
            <xsl:attribute name="href">/xslGip/<xsl:apply-templates select="//MpStyle" />/css/lp_g.css</xsl:attribute>
        </link>
        <xsl:value-of disable-output-escaping="yes" select="../HeaderPart" />
        <table class="Thumbnail">
            <tr>
                <xsl:for-each select="Article">
                    <td>
                        <a target="_blank">
                            <xsl:attribute name="href">
                                <xsl:value-of select="xImgFile" />
                            </xsl:attribute>
                            <xsl:attribute name="title">
                                <xsl:value-of select="ArticleField[fieldName='sTitle']/Value" />
                            </xsl:attribute>
                            <xsl:if test="@newWindow='Y'">
                                <xsl:attribute name="target">_nwGip</xsl:attribute>
                            </xsl:if>
                            <img border="0">
                                <xsl:attribute name="src">
                                    <xsl:value-of select="xImgFile" />
                                </xsl:attribute>
                                <xsl:attribute name="alt">
                                    <xsl:value-of select="ArticleField[fieldName='sTitle']/Value" />
                                </xsl:attribute>
                            </img>
                        </a>
                        <h5>
                            <a>
                                <xsl:attribute name="href">
                                    <xsl:value-of select="ArticleField[fieldName='sTitle']/xURL" />
                                </xsl:attribute>
                                <xsl:attribute name="title">
                                    <xsl:value-of select="ArticleField[fieldName='sTitle']/Value" />
                                </xsl:attribute>
                                <xsl:if test="@newWindow='Y'">
                                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                                </xsl:if>
                                <xsl:value-of select="ArticleField[fieldName='sTitle']/Value" />
                            </a>
                        </h5>
                        <p>
                            <xsl:value-of select="Abstract" />
                        </p>
                    </td>
                    <xsl:if test="position() mod 3 = 0">
                        <tr />
                    </xsl:if>
                </xsl:for-each>
            </tr>
        </table>
        <!--
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
        -->
        <xsl:value-of disable-output-escaping="yes" select="../FooterPart" />
    </xsl:template>
    <!-- 主資料區／結束-->
</xsl:stylesheet>
