<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
	<xsl:include href="headpart.xsl" />
	<xsl:include href="hySearch.xsl" />
	<xsl:include href="epaper.xsl"/>
	<xsl:include href="footer.xsl"/>
	<xsl:include href="menuitem.xsl"/>
	<xsl:include href="info.xsl" />
	
	<xsl:template match="hpMain">
		<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
			<xsl:apply-templates select="." mode="WInfo" />

			<body class="body">
		<!-- 網頁標題 Start-->
					<xsl:apply-templates select="." mode="headinfo" />
		<!-- 網頁標題 End-->

		
	<div id="middle">
	<table class="layout" summary="排版表格">	
		<tr>
		<!-- leftcol start -->	
			<td class="leftbg">
			<div id="leftcol"> 
	
			<!-- accesskey start -->
			<div class="accesskey"><a href="accesskey.htm" title="左方區塊" accesskey="L">:::</a></div>
			<!-- accesskey start -->	
	

			<!-- leftbutton start -->
			<!-- 主要選單 Start-->
				<div class="menu">
					<xsl:for-each select="MenuBar">
						<xsl:apply-templates select="MenuCat" />
					</xsl:for-each>				
				</div>
			</div>
			</td>
			<!-- 主要選單 End-->
		<!-- 左方功能選單區 End-->

		<!-- 中央主資料區 Start-->
			<td id="center">
			<!-- accesskey start -->
			<div class="accesskey"><a href="accesskey.htm" title="中央區塊" accesskey="C">:::</a></div>
			<!-- accesskey start -->	

			<!-- friendly Start -->
			<div class="path">
				<p><xsl:apply-templates select="." mode="xPath"/></p>
			</div>
				<h4><xsl:value-of select="CtUnitName" /></h4>
				<p>本網站依無障礙網頁設計原則建置，網站的主要內容分為四大區塊：<br/>
 					1. 上方功能區塊、2. 左方導覽區塊、3. 中央內容區塊。</p>
				<p>本網站的定位點﹝Accesskey﹞設定如下：</p>
				<p>
					Alt+U：上方功能區塊，包括回首頁、網站導覽等。<br/>
					Alt+L：左方導覽區塊，為本網站的主選單和相關連結區。<br/>
					Alt+C：中央內容區塊，為本頁主要內容區。<br/>
				</p>
				<div class="sitemap">
					<xsl:apply-templates select="." mode="Sitemap"/>
				</div>	
	
			<div class="top"><a><xsl:attribute name="href">dp.asp?mp=<xsl:value-of select="mp" /></xsl:attribute>首頁</a></div>
			</td>
			<!-- content end -->
			</tr>
	</table>
	</div>
		<!-- 中央主資料區 End-->

		
		
		<!-- 頁尾資訊 Start-->
				<xsl:apply-templates select="." mode="Copyright" />
		<!-- 頁尾資訊 End-->
			</body>
		</html>
	</xsl:template>


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
