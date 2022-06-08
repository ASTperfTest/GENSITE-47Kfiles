<?xml version="1.0"  encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:user="urn:user-namespace-here" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone= "yes"/>

<xsl:template match="ePaper">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><xsl:value-of select="ePaperTitle" /></title>
<link rel="stylesheet" type="text/css">
	<xsl:attribute name="href"><xsl:value-of select="ePaperXMLCSSURL" /></xsl:attribute>
</link>
</head>
<body>

<div id="epaperall">

<div id="epapertop">
<div id="epapertime">�o���� : <xsl:value-of select="ePubDate" /></div>
</div>
	<xsl:apply-templates select="epSectionList"/>


<div id="copyright">	
	��F�|�s�D�� <a href="http://www.gio.gov.tw" target="_blank">http://www.gio.gov.tw</a>  <br />
  	�Y�z�������ĳ�A�w��z�ӫH�P�ڭ̤��ɡI�A�ȫH�c�G<a href="mailto:service@gio.gov.tw">service@gio.gov.tw</a>  <br />
  	���v�Ҧ��G��F�|�s�D���@ �x�_���Ѭz��2���@�`���G(02)3356-8888 (�N��) 
  	<br />
	<a>
		<xsl:attribute name="href"><xsl:value-of select="ePaperOrderCancelURL" /></xsl:attribute>
		<xsl:attribute name="title">����/�q�\�q�l��</xsl:attribute>
		<xsl:attribute name="target">_blank</xsl:attribute>
		����/�q�\�q�l��
	</a>  	
</div>

</div>
</body>
</html>

</xsl:template>

<xsl:template match="epSectionList">
	<div id="content">
		<xsl:for-each select="epSection">
		    <xsl:choose>
          		<xsl:when test="position()=1">		        	
				<xsl:apply-templates select="." mode="epSection1" />
			</xsl:when>
	  		<xsl:otherwise>		        
				<xsl:apply-templates select="." mode="epSection2" />
	  		</xsl:otherwise>		        
		    </xsl:choose>				
		</xsl:for-each>  
      	</div>
</xsl:template>

<xsl:template match="epSection" mode="epSection1">
	<div id="topic">
		<xsl:apply-templates select="xItemList" mode="xItemList1" />
	</div>
</xsl:template>

<xsl:template match="xItemList" mode="xItemList1">
	<xsl:if test="position()=1">
		<xsl:if test="xImgFile">
			<p>
			  <img width="200" height="112" hspace="10" border="1" align="left">
				<xsl:attribute name="src"><xsl:value-of select="xItemURL" /><xsl:value-of select="xImgFile" /></xsl:attribute>
				<xsl:attribute name="alt"><xsl:value-of select="sTitle" /></xsl:attribute>
			  </img>
			</p>
		</xsl:if>	
		<p class="contenttitle"><xsl:value-of select="sTitle" /></p>
		<p><xsl:value-of disable-output-escaping="yes" select="xBody" />
 		<a>
		<xsl:attribute name="href"><xsl:value-of select="xItemURL" />/ct.asp?xItem=<xsl:value-of select="xItem" />&amp;ctNode=<xsl:value-of select="../secID" /></xsl:attribute>
		<xsl:attribute name="title"><xsl:value-of select="sTitle" /></xsl:attribute>
		<xsl:attribute name="target">_blank</xsl:attribute>
		(�ԥ���)
		</a>      		
		</p>
	</xsl:if>
</xsl:template>

<xsl:template match="epSection" mode="epSection2">
	<div id="topic">
	<div id="topictitle"><xsl:value-of select="secName" /></div>
	<ul>
		<xsl:for-each select="xItemList">
			<xsl:apply-templates select="." mode="xItemList2" />
		</xsl:for-each>  
	</ul>			
	<p align="right">
	<a>
		<xsl:attribute name="href"><xsl:value-of select="secURL" />/np.asp?ctNode=<xsl:value-of select="secID" /></xsl:attribute>
		<xsl:attribute name="title">��h�T��</xsl:attribute>
		<xsl:attribute name="target">_blank</xsl:attribute>
		��h&gt;&gt;
	</a>
	</p>
	</div>
</xsl:template>



<xsl:template match="xItemList" mode="xItemList2">
        <li>
 	<a>
		<xsl:attribute name="href"><xsl:value-of select="xItemURL" />/ct.asp?xItem=<xsl:value-of select="xItem" />&amp;ctNode=<xsl:value-of select="../secID" /></xsl:attribute>
		<xsl:attribute name="title"><xsl:value-of select="sTitle" /></xsl:attribute>
		<xsl:attribute name="target">_blank</xsl:attribute>
		<xsl:value-of select="sTitle" />
	</a>          
        </li>    
</xsl:template>

</xsl:stylesheet>

