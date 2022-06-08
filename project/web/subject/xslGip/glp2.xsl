<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:saxon="http://icl.com/saxon/"
    extension-element-prefixes="saxon" version="1.0">
    <!-- 主資料區／開始-->
    <xsl:template match="TopicList">
        <link rel="stylesheet" type="text/css">
            <xsl:attribute name="href">/xslGip/<xsl:apply-templates select="//MpStyle" />/css/lp_a.css</xsl:attribute>
        </link>
        <xsl:value-of disable-output-escaping="yes" select="../HeaderPart" />
 <xsl:for-each select="Article">
	<div class="First1">
				<h1><xsl:value-of select="ArticleField[fieldName='sTitle']/Value" /></h1>
				<div class="Head">
					<img src="/xslgip/style01/images/zbg_04.gif" alt="*" />
			
				</div>
				<div class="Body">
					<img> 
				 	<xsl:attribute name="src">
                                    	<xsl:value-of select="xImgFile" />
                                	</xsl:attribute>
                                	<xsl:attribute name="alt">
                                    	<xsl:value-of select="ArticleField[fieldName='sTitle']/Value" />
                                	</xsl:attribute>
					</img>
					<!--img src="images/img01.gif" alt="圖文說明"/-->
					<p><xsl:value-of disable-output-escaping="yes" select="ArticleField[fieldName='xBody']/Value" /></p>
				</div>
				<div class="Foot"><img src="/xslgip/style01/images/zbg_08.gif" alt="*"/></div>		
	</div>
       
</xsl:for-each>
 
        <xsl:value-of disable-output-escaping="yes" select="../FooterPart" />
    </xsl:template>
    <!-- 主資料區／結束-->
</xsl:stylesheet>
