<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
	<xsl:include href="headpart.xsl" />
	<xsl:include href="hySearch.xsl" />
	<xsl:include href="epaper.xsl"/>
	<xsl:include href="footer.xsl"/>

	<xsl:template match="hpMain">
		<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
			<head>
				<title>GIP</title>
				<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
				<link href="xslGIP/style1/css/home.css" rel="stylesheet" type="text/css" />
			</head>
			<body>
		<!-- 網頁標題 Start-->
					<xsl:apply-templates select="." mode="headinfo" />
		<!-- 網頁標題 End-->

		<!-- 快速導覽列 Start-->
				<div class="Title"></div>
					<xsl:apply-templates select="MenuBar2" />
		<!-- 快速導覽列 End-->

		<!-- 左方功能選單區 Start-->
				<div class="Func">
				<!-- 主要選單 Start-->
					<xsl:for-each select="MenuBar">
						<xsl:apply-templates select="MenuCat" />
					</xsl:for-each>
				<!-- 主要選單 End-->
				<!-- AD Start-->
					<xsl:apply-templates select="AD" />
				<!-- AD End-->
				</div> <!--class="Func"-->
		<!-- 左方功能選單區 End-->

		<!-- 中央主資料區 Start-->
				<div class="Main">
					<div class="News">
					<xsl:apply-templates select="." mode="xPath"/>
					<hr/>						
						<xsl:apply-templates select="." mode="Sitemap"/>	
					</div>
				</div>
		<!-- 中央主資料區 End-->

		<!-- 右方資料區 Start-->
				<div class="Column">
					<xsl:apply-templates select="." mode="xsSearch" />
					<xsl:apply-templates select="." mode="xsEPaper" />
				</div>
		<!-- 右方資料區 End-->
		
		<!-- 頁尾資訊 Start-->
				<xsl:apply-templates select="Copyright" />
		<!-- 頁尾資訊 End-->
			</body>
		</html>
	</xsl:template>

<!-- 快速導覽列 -->
	<xsl:template match="MenuBar2">
		<div class="Quickview">
		<xsl:for-each select="MenuCat">
			│<a>
					<xsl:attribute name="href">
						<xsl:value-of select="redirectURL" />
					</xsl:attribute>
					<xsl:if test="@newWindow='Y'">
						<xsl:attribute name="target">_nwGip</xsl:attribute>
					</xsl:if>
					<xsl:value-of select="Caption" />
				</a>		
		</xsl:for-each>
	│</div>
	</xsl:template>
<!-- 快速導覽列 -->

<!-- 主要選單-->
	<xsl:template match="MenuCat">
		<div class="Link">
			<div class="Head">
				<a>
					<xsl:attribute name="href">
						<xsl:value-of select="redirectURL" />
					</xsl:attribute>
					<xsl:if test="@newWindow='Y'">
						<xsl:attribute name="target">_nwGip</xsl:attribute>
					</xsl:if>
					<xsl:value-of select="Caption" />
				</a>
			</div>
			<xsl:if test="MenuItem">
				<div class="Body">
					<xsl:for-each select="MenuItem">
						<xsl:apply-templates select="." />
					</xsl:for-each>
				</div>
			</xsl:if>
		</div>
	</xsl:template>

	<xsl:template match="MenuItem">
		<a>
			<xsl:attribute name="href">
				<xsl:value-of select="redirectURL" />
			</xsl:attribute>
			<xsl:if test="@newWindow='Y'">
				<xsl:attribute name="target">_nwGip</xsl:attribute>
			</xsl:if>
			<xsl:value-of select="Caption" />
		</a>
	</xsl:template>
<!-- 主要選單-->

<!--  路徑連結  -->
	<xsl:template match="hpMain" mode="xPath">
			<a href="default.asp">首頁</a>&gt;網站導覽
<!--  路徑連結  -->
	</xsl:template>  
	
<!-- 網站導覽 -->
<xsl:template match="hpMain" mode="Sitemap">
	<ul> 
	<xsl:for-each select="Sitemap/Menucat">
	 <xsl:if test="@dataparent='0'">
          <li>
          <xsl:choose>	
          <xsl:when test="xURL!=''">
                <a>
		  <xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
		  <xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute> 
		  <xsl:value-of select="Caption" />     	        		
    	        </a>	
	  </xsl:when>
	  <xsl:otherwise>
	          <xsl:value-of select="Caption" />   
	  </xsl:otherwise>
          </xsl:choose>	
          </li> 
          <xsl:variable name="thisCaptionID">
            <xsl:value-of select="@iCuItem" />
          </xsl:variable>
          <xsl:if test="@isUnit='C'">
          	<xsl:for-each select="ancestor::hpMain//Sitemap/Menucat">
             	<xsl:if test="@dataparent=$thisCaptionID">
             	<ul> 
            	<li>
            	<xsl:choose>	
          <xsl:when test="xURL!=''">
                <a>
		  <xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
		  <xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute> 
		  <xsl:value-of select="Caption" />     	        		
    	        </a>	
	  </xsl:when>
	  <xsl:otherwise>
	          <xsl:value-of select="Caption" />   
	  </xsl:otherwise>
          </xsl:choose>		
            	</li> 
            	<xsl:variable name="thisCaptionID2">
            	<xsl:value-of select="@iCuItem" />
         	</xsl:variable>
         	<xsl:if test="@isUnit='C'">
          	<xsl:for-each select="ancestor::hpMain//Sitemap/Menucat">
          	<xsl:if test="@dataparent=$thisCaptionID2">
           	<ul> 
              	<li>
              	<xsl:choose>	
          <xsl:when test="xURL!=''">
                <a>
		  <xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
		  <xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute> 
		  <xsl:value-of select="Caption" />     	        		
    	        </a>	
	  </xsl:when>
	  <xsl:otherwise>
	          <xsl:value-of select="Caption" />   
	  </xsl:otherwise>
          </xsl:choose>	
              	</li> 
              	<xsl:variable name="thisCaptionID3">
            	<xsl:value-of select="@iCuItem" />
         	</xsl:variable>
         	<xsl:if test="@isUnit='C'">
         	<xsl:for-each select="ancestor::hpMain//Sitemap/Menucat">
         	<xsl:if test="@dataparent=$thisCaptionID3">
              	<ul> 
                <li>
               <xsl:choose>	
          <xsl:when test="xURL!=''">
                <a>
		  <xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
		  <xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute> 
		  <xsl:value-of select="Caption" />     	        		
    	        </a>	
	  </xsl:when>
	  <xsl:otherwise>
	          <xsl:value-of select="Caption" />   
	  </xsl:otherwise>
          </xsl:choose>	
                </li> 
                <xsl:variable name="thisCaptionID4">
            	<xsl:value-of select="@iCuItem" />
         	</xsl:variable>
         	<xsl:if test="@isUnit='C'">
         	<xsl:for-each select="ancestor::hpMain//Sitemap/Menucat">
         	<xsl:if test="@dataparent=$thisCaptionID4">
              	<ul> 
                <li>
               <xsl:choose>	
          <xsl:when test="xURL!=''">
                <a>
		  <xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
		  <xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute> 
		  <xsl:value-of select="Caption" />     	        		
    	        </a>	
	  </xsl:when>
	  <xsl:otherwise>
	          <xsl:value-of select="Caption" />   
	  </xsl:otherwise>
          </xsl:choose>	
                </li> 
              	</ul> 
              	</xsl:if>
                </xsl:for-each>
                </xsl:if>
              	</ul> 
              	</xsl:if>
                </xsl:for-each>
                </xsl:if>
                </ul>
                </xsl:if>
            </xsl:for-each>
          	</xsl:if>
          	</ul> 
          	</xsl:if>
          	</xsl:for-each>
             
           </xsl:if>
          </xsl:if> 
         </xsl:for-each>
        </ul> 
</xsl:template>	 
<!-- 網站導覽 -->
</xsl:stylesheet>
