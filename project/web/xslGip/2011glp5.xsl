<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
xmlns:user="urn:user-namespace-here" version="1.0" >
  <!-- 條列／開始 -->
  <xsl:template match="TopicList">
    <link rel="stylesheet" type="text/css" href="/css/jquery.autocomplete.css" />
    <div class="subjectlist">
      <ul>
        <xsl:for-each select="Article">
          <li>
            <a>
              <xsl:attribute name="href">
                <xsl:value-of disable-output-escaping="yes" select="ArticleField/xURL" />
              </xsl:attribute>
              <xsl:attribute name="title">
                <xsl:value-of select="ArticleField[fieldName='stitle']/Value" />
              </xsl:attribute>
              <xsl:if test="ArticleField/xURL[@newWindow='Y']">
                <xsl:attribute name="target">_blank</xsl:attribute>
              </xsl:if>
              <xsl:value-of select="ArticleField[fieldName='stitle']/Value" />
              <img>
                <xsl:attribute name="style">width:160px;height:auto;</xsl:attribute>
                
                  <xsl:choose>
                    <xsl:when test="xImgFile !=''">
                      <xsl:attribute name="src"><xsl:value-of disable-output-escaping="yes" select="xImgFile" />
                      </xsl:attribute>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:if test="xRandom = 1">
                          <xsl:attribute name="src">images/2011glp1.gif</xsl:attribute>
                        </xsl:if>
                        <xsl:if test="xRandom = 2">
                          <xsl:attribute name="src">images/2011glp2.gif</xsl:attribute>
                        </xsl:if>
                        <xsl:if test="xRandom = 3">
                          <xsl:attribute name="src">images/2011glp3.gif</xsl:attribute>
                        </xsl:if>
                        <xsl:if test="xRandom = 4">
                          <xsl:attribute name="src">images/2011glp4.gif</xsl:attribute>
                        </xsl:if>

                    </xsl:otherwise>
                  </xsl:choose>
                
                <xsl:attribute name="alt">
                  <xsl:value-of disable-output-escaping="yes" select="ArticleField[fieldName='stitle']/Value" />
                </xsl:attribute>
                <xsl:attribute name="title">
                  <xsl:value-of disable-output-escaping="yes" select="ArticleField[fieldName='stitle']/Value" />
                </xsl:attribute>
              </img>
            </a>
            <xsl:value-of disable-output-escaping="yes" select="ArticleField[fieldName='xabstract']/Value" />

            <xsl:if test="xURL !=''">
              <a>
                <xsl:attribute name="href">
                  <xsl:value-of select="xURL" />
                </xsl:attribute>
                <xsl:attribute name="title">
                  <xsl:value-of select="Caption" />
                </xsl:attribute>
                <xsl:if test="@newWindow='Y'">
                  <xsl:attribute name="target">_blank</xsl:attribute>
                </xsl:if>
                <xsl:value-of select="Caption" />
              </a>
            </xsl:if>
            <xsl:if test="xURL =''">
              <xsl:value-of select="Caption" />
            </xsl:if>
            &#x9;
            <xsl:value-of select="DeptName" />
            &#x9;
            <xsl:value-of select="PostDate" />
            <span>
              <xsl:value-of select="body" disable-output-escaping="yes" />
            </span>
          </li><hr/>
        </xsl:for-each>
      </ul>
      <!--div class="more"><a href="#">更多訊息</a></div-->
    </div>
  </xsl:template>
  <!-- 條列／結束 -->
</xsl:stylesheet>

