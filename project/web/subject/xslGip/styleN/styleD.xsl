<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />

<xsl:template match="MainArticle">
	<div id="cp4">
			<h4><xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='stitle']/Value" /></h4>
      <xsl:if test="not(MainArticleField[fieldName='fileDownLoad']/Value='')">
		    <div class="video">
			    <embed width="450" height="340" autostart="ture">
            <xsl:attribute name="src">
              /subject/public/data/<xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='fileDownLoad']/Value"/>
            </xsl:attribute>
			    </embed>
		    </div>
      </xsl:if>
        <table summary="排版表格">
            <!--
                <tr>
                    <th scope="row">標題</th>
                    <td>
                        <xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
                    </td>
                </tr>
                -->
            <xsl:if test="not(MainArticleField[fieldName='xbody']/Value='')">
                <tr>
                    <th scope="row">內文</th>
                    <td width="100%">
                        <SPAN class="idx2" id="article" name="article">
                            <xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xbody']/Value" />
                        </SPAN>
                    </td>
                </tr>
            </xsl:if>
        </table>
	</div>
</xsl:template>

<xsl:template match="AttachmentList" mode="listx">
</xsl:template>


<!--xsl:template match="MainArticleField">    
	<th scope="row"><xsl:value-of select="Title" /></th>
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
					<xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="../MainArticleField[fieldName='xurl']/Value" /></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="Value" /></xsl:attribute> 
					<xsl:if test="xURL[@newWindow='Y']">
						<xsl:attribute name="target">_blank</xsl:attribute>
					</xsl:if>
					<xsl:value-of select="Value" />       	        		
				</a>			
			</td>
		</xsl:when>
		<xsl:otherwise><td><xsl:value-of disable-output-escaping="yes" select="Value" /></td></xsl:otherwise>
	</xsl:choose>
</xsl:template-->   


</xsl:stylesheet>
