<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:user="urn:user-namespace-here" version="1.0">

<xsl:template match="TopicList">
	<xsl:value-of disable-output-escaping="yes" select="../HeaderPart" />
		<table border="0" align="center" cellpadding="0" cellspacing="0" xmlns="">
			<caption>�w��ϥΰݨ��լd</caption>
			 <tbody>
				<xsl:apply-templates select="../TopicTitleList" />		
				<xsl:for-each select="Article">
					<tr bgcolor="#E7E7E7">
						<td><xsl:value-of select="position()"/></td>
						<xsl:apply-templates select="."/>
					</tr>
				</xsl:for-each>
			</tbody>
  		</table>
	<xsl:value-of disable-output-escaping="yes" select="../FooterPart" />
</xsl:template>
    
<xsl:template match="TopicTitleList">    
   	<tr bgcolor="#BDDAE8">
  	<th>�s��</th>
  	<th>�D�D</th>
  	<th>�լd�_�W�ɶ�</th>
  	<th>�ѻP�լd</th>
  	<th>�[�ݵ��G</th>
	</tr>
</xsl:template>
   
<xsl:template match="TopicList/Article">    
	<td><xsl:apply-templates select="ArticleField[fieldName='sTitle']"/></td>
	<td><xsl:apply-templates select="ArticleField[fieldName='m011_bdate']"/>~<xsl:apply-templates select="ArticleField[fieldName='m011_edate']"/></td>
	<td nowarp="true">
		<!--xsl:if test="@xInDateRange='Y'"-->
		<a>
			<xsl:if test="ArticleField[fieldName='m011_jumpquestion']/Value='0'">
				<xsl:attribute name="href">ap.asp?xdURL=sVote/vote01.asp&amp;subjectId=<xsl:value-of select="@iCuItem" />&amp;ctNode=<xsl:value-of select="//@xNode" /></xsl:attribute>
	  		</xsl:if>
			<xsl:if test="ArticleField[fieldName='m011_jumpquestion']/Value='1'">
				<xsl:attribute name="href">ap.asp?xdURL=sVote/vote01_jump.asp&amp;subjectId=<xsl:value-of select="@iCuItem" />&amp;ctNode=<xsl:value-of select="//@xNode" /></xsl:attribute>
	  		</xsl:if>	  		
	  		<img src="images/Service/vote.gif" alt="�ѻP�լd" width="25" height="25" hspace="5" border="0" align="absmiddle"/>�ѻP�լd
		</a>
		<!--/xsl:if-->
	&amp;nbsp;</td>
	<td nowarp="true">
		<a>
	  		<xsl:attribute name="href">ap.asp?xdURL=sVote/vote02.asp&amp;subjectId=<xsl:value-of select="@iCuItem" />&amp;ctNode=<xsl:value-of select="//@xNode" /></xsl:attribute>
	  		<img src="images/Service/result.gif" alt="�[�ݵ��G" width="25" height="25" hspace="5" border="0" align="absmiddle"/>�[�ݵ��G
		</a>
	</td>
</xsl:template>

<xsl:template match="ArticleField">    
	<xsl:choose>	
		<xsl:when test="xURL">
    	        		<a>
		  		<xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
		  		<xsl:attribute name="title"><xsl:value-of select="Value" /></xsl:attribute> 
				<xsl:if test="xURL[@newWindow='Y']">
					<xsl:attribute name="target">_nwGip</xsl:attribute>
				</xsl:if>
		  		<xsl:value-of select="Value" />       	        		
    	        		</a>			
		</xsl:when>
		<xsl:otherwise>
  				<xsl:value-of select="Value" />    	        				
		</xsl:otherwise>
	</xsl:choose>	
</xsl:template>

</xsl:stylesheet>

