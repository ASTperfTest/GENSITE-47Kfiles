<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
xmlns:user="urn:user-namespace-here" version="1.0">
    <xsl:template match="MainArticle">



        <div id="cp">
            <div class="experts">
                <h4>
                    <xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='stitle']/Value" />
                </h4>
                <div class="video">
                    <!--embed width="450" height="340" autostart="ture"-->

                    <OBJECT ID="MediaPlayer" WIDTH="450" HEIGHT="340" CLASSID="CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95"
                    STANDBY="Loading Windows Media Player components..." TYPE="application/x-oleobject">
                        <PARAM NAME="FileName">
                            <xsl:choose>
                                <xsl:when test="//AttachmentList/Attachment/Descxx!=''">
                                    <xsl:attribute name="VALUE">
                                        <xsl:value-of disable-output-escaping="yes" select="//AttachmentList/Attachment/Descxx"/>
                                    </xsl:attribute>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:attribute name="VALUE">
                                        http://kmms.coa.gov.tw/<xsl:value-of disable-output-escaping="yes" select="//AttachmentList/Attachment/URL"/>
                                    </xsl:attribute>
                                </xsl:otherwise>
                            </xsl:choose>
                        </PARAM>
                        <PARAM name="ShowControls" VALUE="true"/>
                        <param name="ShowStatusBar" value="true"/>
                        <PARAM name="ShowDisplay" VALUE="false"/>
                        <PARAM name="autostart" VALUE="true"/>
                        <EMBED TYPE="application/x-mplayer2"
                        NAME="MediaPlayer"
                        WIDTH="450" HEIGHT="340" ShowControls="1" ShowStatusBar="1" ShowDisplay="0" autostart="1">
                            <xsl:choose>
                                <xsl:when test="//AttachmentList/Attachment/Descxx!=''">
                                    <xsl:attribute name="SRC">
                                        <xsl:value-of disable-output-escaping="yes" select="//AttachmentList/Attachment/Descxx"/>
                                    </xsl:attribute>
                                </xsl:when>
                                <xsl:otherwise>
                                    <xsl:attribute name="SRC">
                                        http://kmms.coa.gov.tw/<xsl:value-of disable-output-escaping="yes" select="//AttachmentList/Attachment/URL"/>
                                    </xsl:attribute>
                                </xsl:otherwise>
                            </xsl:choose>
                        </EMBED>
                    </OBJECT>
                </div>
                <span id="article" name="article" class="idx1">
                    <table summary="排版表格" class="cptable">
                        <col class="title" />
                        <col class="cptablecol2" />
                        <xsl:for-each select="MainArticleField">
                            <tr>
                                <xsl:apply-templates select="."/>
                            </tr>
                        </xsl:for-each>
                    </table>
                </span>
                <!--p><xsl:value-of disable-output-escaping="yes" select="//MainArticleField[fieldName='xbody']/Value" /></p-->
            </div>
        </div>

        <!--div id="cp">
		<div class="experts">
			<h4><xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='stitle']/Value" /></h4>
			<div class="video">
				<embed width="450" height="320" autostart="false">
				<xsl:attribute name="src">public/MMO/video/<xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='mmoFileName']/Value"/></xsl:attribute>
				</embed>
			</div>
			<p><xsl:value-of disable-output-escaping="yes" select="//MainArticleField[fieldName='xbody']/Value" /></p>
		</div>
	</div-->
    </xsl:template>

    <xsl:template match="MainArticleField">
        <th scope="row" nowrap="nowrap">
            <xsl:value-of select="Title" />
        </th>
        <xsl:choose>
            <xsl:when test="fieldName='ximgFile'">
                <td>
                    <img>
                        <xsl:attribute name="src">
                            public/Data/<xsl:value-of disable-output-escaping="yes" select="../MainArticleField[fieldName='ximgFile']/Value" />
                        </xsl:attribute>
                        <xsl:attribute name="alt">
                            <xsl:value-of disable-output-escaping="yes" select="../MainArticleField[fieldName='stitle']/Value" />
                        </xsl:attribute>
                        <xsl:attribute name="title">
                            <xsl:value-of disable-output-escaping="yes" select="../MainArticleField[fieldName='stitle']/Value" />
                        </xsl:attribute>
                    </img>
                </td>
            </xsl:when>
            <xsl:when test="fieldName='xurl'">
                <td>
                    <a>
                        <xsl:attribute name="href">
                            <xsl:value-of disable-output-escaping="yes" select="../MainArticleField[fieldName='xurl']/Value" />
                        </xsl:attribute>
                        <xsl:attribute name="title">
                            <xsl:value-of select="Value" />
                        </xsl:attribute>
                        <xsl:if test="xURL[@newWindow='Y']">
                            <xsl:attribute name="target">_blank</xsl:attribute>
                        </xsl:if>
                        <xsl:value-of select="Value" />
                    </a>
                </td>
            </xsl:when>

            <xsl:otherwise>
                <td>
                    <xsl:value-of disable-output-escaping="yes" select="Value" />
                </td>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:stylesheet>

