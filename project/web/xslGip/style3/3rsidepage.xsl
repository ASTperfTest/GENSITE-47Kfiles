<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
<xsl:include href="info.xsl"/>

<xsl:template match="hpMain">
	<html xmlns="http://www.w3.org/1999/xhtml">
		<head>
			<xsl:apply-templates select="." mode="Header"/>
		</head>
		<script language="javascript" type="text/javascript">
			$(document).ready(function() {
			    loadtagcloud();
			});
			function loadtagcloud() {
			var tagcloudUrl="/services/tagCloud.aspx?";
			var params = "cmd=gettagcloud";
			$.ajax({url:tagcloudUrl,data:params,processData:true,async:false,success:function(data){var result = data; writeTagcloud(result);}});
					
         }
		 function writeTagcloud(data)
		 {
			if (data == null) {
            return;
			}
			$("#tagcloud").html(data);
		 }
		</script>
		<body>
			<!--header star (style A)-->
			<xsl:apply-templates select="." mode="RunActive"/>
			<!--header End-->
			<!--nav star (style A)-->
			<xsl:apply-templates select="." mode="toplink"/>
			<!--nav end-->
			<!--subnav star (style G)-->
			<xsl:apply-templates select="." mode="weblink"/>
			<!--subnav end-->

			<!--layout Table star (style A)-->
			<table class="layout">
			<tr>
				<xsl:apply-templates select="pHTML"/>
				
				<td class="right">
					<xsl:apply-templates select="." mode="r"/>
					<!--search star (style C)-->
					<xsl:apply-templates select="." mode="xsSearch"/>
					<!--search end-->
					<!--login star-->
					<xsl:apply-templates select="." mode="xsMember"/>
					<!--login end-->
					<!--epaper star (style B)-->
					<xsl:apply-templates select="." mode="xsePaper"/>
					<!--epaper end-->
					<table>						
						<tr>
							<td>
								<!--最近瀏覽-->
								<xsl:if test="contains(//redirectUrl, 'knowledge')">
									<xsl:apply-templates select="." mode="newest"/>
								</xsl:if>
								<!--最近瀏覽end-->								
							</td>
						</tr>
						<tr>
							<td>
								<!--tag Cloud -->
								<div class="TagCloud">
									<h2>熱門關鍵字</h2>
									<div id ="tagcloud">
									</div>
								</div>
								<!--ag Cloud end-->
							</td>
						</tr>						
					</table>
				</td>
			  </tr>
			</table>
			<!--layout Table End-->
			<!--footer star (style A)-->
			<div class="footer">
				<xsl:apply-templates select="." mode="xsFooter"/>		
				<xsl:apply-templates select="." mode="xsCopyright"/>		
			</div>
			<!--footer End-->
		</body>
	</html>
</xsl:template>

<!-- 中央主資料區 Start-->
<xsl:template match="pHTML">
<xsl:value-of disable-output-escaping="yes" select="." />
<!--xsl:copy-of select="."/-->
</xsl:template>
</xsl:stylesheet>
