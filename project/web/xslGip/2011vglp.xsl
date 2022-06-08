<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
xmlns:user="urn:user-namespace-here" version="1.0">

	
	<!-- 資料大類 -->
<xsl:template match="CatList">
    <div class="category">
<span class="cateName">主分類</span>
<ul>
<li><a><xsl:attribute name="href">lp.asp?<xsl:value-of select="../xqURL" /></xsl:attribute>
        <xsl:if test="not(contains(//qURL, 'xq_xCat='))"><b>All</b></xsl:if>
        <xsl:if test="(contains(//qURL, 'xq_xCat='))">All</xsl:if>
        </a></li>
    		<xsl:for-each select="CatItem">
<li>
    			<a>
    				<xsl:attribute name="href">lp.asp?<xsl:value-of select="../../xqURL" /><xsl:value-of select="xqCondition" /></xsl:attribute>
    				<xsl:if test="contains(concat(//qURL, ''), concat(xqCondition, ''))"><b>
            <xsl:value-of select="CatName" /></b>
            </xsl:if>
            	<xsl:if test="not(contains(concat(//qURL, ''), concat(xqCondition, '')))">
            <xsl:value-of select="CatName" />
            </xsl:if>
    			</a>
</li>
    		</xsl:for-each>
</ul>
    </div>
    		<xsl:apply-templates select="//CatListRelated" />
</xsl:template>
<!-- 資料大類 -->

<xsl:template match="CatListRelated">
<div class="category">
<span class="cateName">次分類</span>
<ul>
  <xsl:for-each select="row">
<li>
      <a>
        <xsl:attribute name="href">lp.asp?<xsl:if test="contains(//qURL, 'htx_vgroup=')"><xsl:value-of select="//xqURLDeleteAnd1"/>&amp;htx_vgroup=<xsl:value-of select="mCode"/>
        </xsl:if>
          
          <xsl:if test="not(contains(//qURL, 'htx_vgroup='))"><xsl:value-of select="//qURL"/>&amp;htx_vgroup=<xsl:value-of select="mCode"/>
          </xsl:if>
        </xsl:attribute>
        <xsl:if test="contains(concat(//qURL, ''), concat(mCode, ''))"><b>
        <xsl:value-of select="mValue" /></b>
         </xsl:if>
         
           <xsl:if test="not(contains(concat(//qURL, ''), concat(mCode, '')))">
        <xsl:value-of select="mValue" />
         </xsl:if>
         
      </a>
</li>
    </xsl:for-each>
</ul>
</div> 
</xsl:template>


    <!-- 條列／開始 -->
    <xsl:template match="TopicList">
        <div class="VideoAlbum">
            <ul class="VideoAlbum">
                <xsl:for-each select="Article">
                    <li>
                        <a>
                            <xsl:attribute name="href">
                                <xsl:value-of disable-output-escaping="yes" select="ArticleField/xURL" />
                            </xsl:attribute>

                            <img>
                                <xsl:attribute name="src">
                                    <xsl:choose>
                                        <xsl:when test="xImgFile !=''">
                                            <xsl:value-of disable-output-escaping="yes" select="xImgFile" />
                                        </xsl:when>
                                        <xsl:otherwise>
                                            images/2011vglp<xsl:value-of select="TopCat" />.jpg
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:attribute>
                                <xsl:attribute name="alt">
                                    <xsl:value-of disable-output-escaping="yes" select="ArticleField[fieldName='stitle']/Value" />
                                </xsl:attribute>
                                <xsl:attribute name="title">
                                    <xsl:value-of disable-output-escaping="yes" select="ArticleField[fieldName='stitle']/Value" />
                                </xsl:attribute>
                            </img>
                            <div style="text-align:center;Height:30px;overflow:hidden;margin:10px">
                                <xsl:attribute name="title">
                                    <xsl:value-of select="ArticleField[fieldName='stitle']/Value" />
                                </xsl:attribute>
                                <xsl:if test="ArticleField/xURL[@newWindow='Y']">
                                    <xsl:attribute name="target">_blank</xsl:attribute>
                                </xsl:if>
                                <xsl:value-of select="Caption" />
                            </div>
                        </a>
                        <!--<xsl:value-of disable-output-escaping="yes" select="ArticleField[fieldName='xabstract']/Value" />-->
                    </li>
                </xsl:for-each>
            </ul>
            <!--div class="more"><a href="#">更多訊息</a></div-->
        </div>
    </xsl:template>
    <!-- 條列／結束 -->
</xsl:stylesheet>

