<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
<!--Start-->
	<xsl:template match="hpMain" mode="WInfo">
			<head>
				<title><xsl:value-of select="rootname" /></title>
				<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />						
      	<link rel="stylesheet" type="text/css">
      		<xsl:attribute name="href">xslGIP/<xsl:value-of select="MpStyle"/>/css/style.css</xsl:attribute>
      	</link>
		  <link rel="stylesheet" type="text/css">
		  	<xsl:attribute name="href">xslgip/<xsl:value-of select="//MpStyle" />/css/login.css</xsl:attribute>
      </link >
	  <link rel="stylesheet" type="text/css">
		                       <xsl:attribute name="href">/css/pedia.css</xsl:attribute>
	                </link>
			</head>
	</xsl:template>
<!--End-->	

</xsl:stylesheet>


