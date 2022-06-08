<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">	
	
<!-- 附件列表 start-->
	<xsl:template match="AttachmentList" mode="listx">
		<h4>附件列表</h4>
		<ul>
			<xsl:for-each select="Attachment">
				<xsl:apply-templates select="." />
			</xsl:for-each>
		</ul>
	</xsl:template>

	<xsl:template match="Attachment">
		<li>		
			<a target="_nwGIP">
				<xsl:attribute name="href">
					<xsl:value-of select="URL" />
				</xsl:attribute>
				<xsl:attribute name="title">
					<xsl:value-of select="Caption" />
				</xsl:attribute>
				<xsl:value-of select="Caption" />
			</a>
		</li>
	</xsl:template>
<!-- 附件列表 End-->

<!-- 附件圖片列表 start-->


	<xsl:template match="AttachmentList" mode="photox">
		<!--div class="h4" >相關照片</div-->
		<h4>相關照片</h4>

		
			<xsl:for-each select="Attachment">
				<div class="lpphoto2" >
					<a target="_nwGIP">
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
						</img>
					</a>	

					<div style="clear: both;"></div>
					<p>

						<a target="_nwGIP">
							<xsl:attribute name="href">
								<xsl:value-of select="URL" />
							</xsl:attribute>
							<xsl:attribute name="title">
								<xsl:value-of select="Caption" />
							</xsl:attribute>
							<xsl:value-of select="Caption" />
						</a>

					</p>
				</div>
			</xsl:for-each>
		
	</xsl:template>

	
<!-- 附件圖片列表 End-->
</xsl:stylesheet>






