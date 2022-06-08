<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />


<!-- 中央主資料區 Start-->
	<xsl:template match="MainArticle">
		<!--  <div class="News"> -->	
		<h4><xsl:value-of select="MainArticleField[fieldName='stitle']/Value" /></h4>	
		
		<xsl:if test="@IsPic ='Y'">							
			<xsl:if test="not(MainArticleField[fieldName='ximgFile']/Value='')">
				<img class="leftimg">
					<xsl:attribute name="src">
						/subject/public/data/<xsl:value-of select="MainArticleField[fieldName='ximgFile']/Value" />
					</xsl:attribute>
					<xsl:attribute name="alt">
						<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
					</xsl:attribute>
				</img>
			</xsl:if>
		</xsl:if>
			<xsl:if test="@IsPostDate ='Y'">
				<p><em>日期:&amp;nbsp;<xsl:value-of select="MainArticleField[fieldName='xpostDate']/Value" /></em></p>
			</xsl:if>
    <SPAN class="idx2" id="article" name="article">
      <xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xbody']/Value" />
    </SPAN>	
		<!--</div>-->
	</xsl:template>
<!-- 中央主資料區 End-->

<xsl:template match="AttachmentList" mode="listx">
<!-- div class="h4" >相關照片</div--> 
<h4>相關照片</h4> 
 <div class="yoxview">
	<xsl:for-each select="Attachment">
		<div class="lpphoto2">
			<xsl:choose>
			  <xsl:when test="IsImageFile='Y'">		  
					<a>
						<xsl:attribute name="href">
							<xsl:value-of select="URL" /> 
						</xsl:attribute>
					<xsl:attribute name="title">
						<xsl:value-of select="Caption" /> 
					</xsl:attribute>		  
						<img>
							<xsl:attribute name="src">
								<xsl:value-of select="URL" /> 
							</xsl:attribute>
							<xsl:attribute name="alt">
								<xsl:value-of select="Caption" /> 
							</xsl:attribute>
							<xsl:attribute name="title">
								<xsl:value-of select="Caption" /> 
							</xsl:attribute>
						</img>
						<xsl:choose>
							<xsl:when test="Descxx!=''">
								<p>
									<xsl:value-of select="Descxx" />
								</p>
							</xsl:when>
							<xsl:otherwise>
								<p>
									<xsl:value-of select="Caption" />
								</p>
							</xsl:otherwise>
						</xsl:choose>
					</a>
			  </xsl:when>
			  <xsl:otherwise>		
			  
					<a target="_nwGIP">
						<xsl:attribute name="href">
							<xsl:value-of select="URL" /> 
						</xsl:attribute>
					<xsl:attribute name="title">
						<xsl:value-of select="Caption" /> 
					</xsl:attribute>		  
						<img src="/public/file.png" style="height:96px;width:96px" border="0" />    
					</a>				
					<div style="clear: both;" /> 
					<p>
						<a>
							<xsl:attribute name="href">
								<xsl:value-of select="URL" /> 
							</xsl:attribute>
							<xsl:attribute name="title">
								<xsl:value-of select="Caption" /> 
							</xsl:attribute>
								<xsl:value-of select="Caption" /> 
						</a>
					</p>
			  </xsl:otherwise>
			</xsl:choose>		
		</div>
	</xsl:for-each>
</div>	
</xsl:template>

</xsl:stylesheet>
