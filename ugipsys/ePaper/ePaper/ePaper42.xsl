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
<div id="epapertime">發行日期 : <xsl:value-of select="ePubDate" /></div>
</div>
	<xsl:apply-templates select="epSectionList"/>


<div id="copyright">	
	行政院新聞局 <a href="http://www.gio.gov.tw" target="_blank">http://www.gio.gov.tw</a>  <br />
  	若您有任何建議，歡迎您來信與我們分享！服務信箱：<a href="mailto:service@gio.gov.tw">service@gio.gov.tw</a>  <br />
  	版權所有：行政院新聞局　 台北市天津街2號　總機：(02)3356-8888 (代表號) 
  	<br />
	<a>
		<xsl:attribute name="href"><xsl:value-of select="ePaperOrderCancelURL" /></xsl:attribute>
		<xsl:attribute name="title">取消/訂閱電子報</xsl:attribute>
		<xsl:attribute name="target">_blank</xsl:attribute>
		取消/訂閱電子報
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
		(詳全文)
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
		<xsl:attribute name="title">更多訊息</xsl:attribute>
		<xsl:attribute name="target">_blank</xsl:attribute>
		更多&gt;&gt;
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

