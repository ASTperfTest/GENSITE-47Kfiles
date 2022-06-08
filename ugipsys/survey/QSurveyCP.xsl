<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:hyweb="urn:gip-hyweb-com" xmlns:user="urn:user-namespace-here" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone= "yes"/>

<!--主資料區-->
<xsl:template match="MainArticle">	
	<div class="topic"><xsl:value-of select="MainArticleField[fieldName='sTitle']/Value" /></div>
	<table cellspacing="0" summary="資料表格">
		<caption>歡迎使用問卷調查</caption>
		<tr> 
			<th width="15%">問卷主題:</th>
			<td width="85%"><xsl:value-of select="MainArticleField[fieldName='sTitle']/Value" /></td>
		</tr>
 		<tr> 
			<th>問卷說明:</th>
			<td>
					<xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xBody']/Value" />&amp;nbsp;
			</td>
		</tr>
		<tr> 
			<th>調查時間:</th>
			<td>
					<xsl:apply-templates select="MainArticleField[fieldName='m011_bdate']/Value"/> ~ <xsl:apply-templates select="MainArticleField[fieldName='m011_edate']/Value"/>&amp;nbsp;
			</td>
		</tr>
		<tr> 
			<th>參與調查:</th>
			<td>
				<xsl:if test="@xInDateRange='Y'">
					<a>
						<xsl:if test="MainArticleField[fieldName='m011_jumpquestion']/Value='0'">
							<xsl:attribute name="href">ap.asp?xdURL=sVote/vote01.asp&amp;subjectId=<xsl:value-of select="@iCuItem" />&amp;ctNode=<xsl:value-of select="//@xNode" /></xsl:attribute>
						</xsl:if>
						<xsl:if test="MainArticleField[fieldName='m011_jumpquestion']/Value='1'">
							<xsl:attribute name="href">ap.asp?xdURL=sVote/vote01_jump.asp&amp;subjectId=<xsl:value-of select="@iCuItem" />&amp;ctNode=<xsl:value-of select="//@xNode" /></xsl:attribute>
						</xsl:if>	  		
						<IMG src="images/Service/vote.gif" alt="參與調查" width="25" height="25" hspace="5" border="0" align="absmiddle"/>參與調查
					</a>
				</xsl:if>
				&amp;nbsp;
			</td>
		</tr>
		<tr> 
			<th>觀看結果:</th>
			<td>
		<a>
	  		<xsl:attribute name="href">ap.asp?xdURL=sVote/vote02.asp&amp;subjectId=<xsl:value-of select="@iCuItem" />&amp;ctNode=<xsl:value-of select="//@xNode" /></xsl:attribute>
	  		<IMG src="images/Service/result.gif" alt="觀看結果" width="25" height="25" hspace="5" border="0" align="absmiddle"/>觀看結果</a>&amp;nbsp;
			</td>
		</tr>
	</table> 
</xsl:template>
    
</xsl:stylesheet>
