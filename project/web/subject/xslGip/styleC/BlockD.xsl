<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">	
	
<!-- BlockA Start -->
	<xsl:template match="BlockD">	
		
			<!-- dataremark start -->
			<h4>
				<xsl:value-of select="Caption" />
			</h4>
			<!--dataremark end -->
			
				<xsl:choose>
					<xsl:when test="ShowStyle = 'Style1'">
						<ul>
							<xsl:for-each select="Article" >
								<xsl:apply-templates select="." mode="Style1"/>			
							</xsl:for-each>
						</ul>
						<div class="more">
							<a>
								<xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="@xNode" />&amp;mp=<xsl:value-of select="//hpMain/mp"/></xsl:attribute>
								更多訊息
							</a>
						</div>
					</xsl:when>
					<xsl:when test="ShowStyle = 'Style2'">
						<table class="inline" summary="排版表格">
							<tr>
								<xsl:for-each select="Article[position()&lt;5]" >
									<xsl:apply-templates select="." mode="Style2"/>										
								</xsl:for-each>
							</tr>
							
							<xsl:if test="Article[position()&gt;4]">
							<tr>
								<xsl:for-each select="Article[position()&lt;9 and position()&gt;4]" >
									<xsl:apply-templates select="." mode="Style2"/>										
								</xsl:for-each>
							</tr>
							</xsl:if>
						</table>
						<div class="more">
							<a>
								<xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="@xNode" />&amp;mp=<xsl:value-of select="//hpMain/mp"/></xsl:attribute>
								更多訊息
							</a>
						</div>
					</xsl:when>
					<xsl:when test="ShowStyle = 'Style3'">
						<xsl:for-each select="Article[position()=1]" >				
							<xsl:apply-templates select="." mode="Style3"/>		
						</xsl:for-each>			
					</xsl:when>							
					<xsl:otherwise>
					</xsl:otherwise>
				</xsl:choose>
		
	</xsl:template>
<!-- BlockA End -->

		
	<!--文章內容 此為style1的樣式 start -->
	<xsl:template match="//BlockD/Article" mode="Style1">
		<xsl:if test="../IsPic = 'Y'">
			<div class="block">		
				<xsl:if test="xImgFile">
					<img class="listimg">
						<xsl:attribute name="src">
							<xsl:value-of select="xImgFile" />
						</xsl:attribute>
						<xsl:attribute name="alt">
							<xsl:value-of select="Caption" />
						</xsl:attribute>
					</img>
				</xsl:if>
				<h5>
				<a>
					<xsl:attribute name="href">
						<xsl:value-of select="xURL" />
					</xsl:attribute>
					<xsl:attribute name="title">
						<xsl:value-of select="Caption" />
					</xsl:attribute>
					<xsl:if test="@newWindow='Y'">
						<xsl:attribute name="target">_nwGip</xsl:attribute>
					</xsl:if>					
						<xsl:value-of select="Caption" />					
				</a>
					
					<xsl:if test="../IsPostDate = 'Y'">		
						-- <xsl:value-of select="PostDate" />	
					</xsl:if>
				</h5>
			</div>
		</xsl:if>
		<xsl:if test="../IsPic = 'N'">
		<li>
			<a>
				<xsl:attribute name="href">
					<xsl:value-of select="xURL" />
				</xsl:attribute>
				<xsl:attribute name="title">
					<xsl:value-of select="Caption" />
				</xsl:attribute>
				<xsl:if test="@newWindow='Y'">
					<xsl:attribute name="target">_nwGip</xsl:attribute>
				</xsl:if>				
					<xsl:value-of select="Caption" />				
			</a>	
			<xsl:if test="../IsPostDate = 'Y'">		
			-- <xsl:value-of select="PostDate" />	
			</xsl:if>
		</li>
		</xsl:if>
	</xsl:template>
	<!--文章內容 此為style1的樣式 end -->
	
	<!--文章內容 此為style2的樣式 start -->
	<xsl:template match="//BlockD/Article" mode="Style2">
		<td>
			<xsl:if test="../IsPic = 'Y'">		
				<xsl:if test="xImgFile">
					<img class="listimg">
						<xsl:attribute name="src">
							<xsl:value-of select="xImgFile" />
						</xsl:attribute>
						<xsl:attribute name="alt">
							<xsl:value-of select="Caption" />
						</xsl:attribute>
					</img>
				</xsl:if>
			</xsl:if>
			<p>
				<a>
					<xsl:attribute name="href">
						<xsl:value-of select="xURL" />
					</xsl:attribute>
					<xsl:attribute name="title">
						<xsl:value-of select="Caption" />
					</xsl:attribute>
					<xsl:if test="@newWindow='Y'">
						<xsl:attribute name="target">_nwGip</xsl:attribute>
					</xsl:if>					
						<xsl:value-of select="Caption" />					
				</a>
				<br />
				<xsl:if test="../IsPostDate = 'Y'">			
					<xsl:value-of select="PostDate" />	
				</xsl:if>
			</p>
		</td>		
	</xsl:template>
	<!--文章內容 此為style2的樣式 end -->
		<!-- 文章內容  此為style3的樣式 start -->
	<xsl:template match="//BlockD/Article" mode="Style3">
		
			<!-- 文章pic start -->			
		 <xsl:if test="../IsPic = 'Y'">		
			<xsl:if test="xImgFile">
				<img class="headline">
					<xsl:attribute name="src">
						<xsl:value-of select="xImgFile" />
					</xsl:attribute>
					<xsl:attribute name="alt">
						<xsl:value-of select="Caption" />
					</xsl:attribute>
				</img>
			</xsl:if>
		</xsl:if>
			<!-- 文章pic end -->					
			<!--文章標題 start -->
			<h5>			
				<a>
				<!--xsl:attribute name="href">ct.asp?xItem=<xsl:value-of select="@iCuItem" />&amp;ctNode=<xsl:value-of select="../@xNode" /></xsl:attribute-->
						<xsl:attribute name="href">
						<xsl:value-of select="xURL" />
					</xsl:attribute>
          <xsl:value-of select="Caption" />
				</a>
			<!--文章標題 end-->			
			<!-- 預計放張貼日期 start-->
			<xsl:if test="../IsPostDate = 'Y'">
				 -- <xsl:value-of select="PostDate" />
			</xsl:if>
			</h5>
			<!-- 預計放張貼日期 end -->			
			
			<!-- 文章內容 start-->
			<p>
				<xsl:if test="../IsExcerpt ='Y'">				
				<xsl:value-of disable-output-escaping="yes" select="Content" />				
			<!-- 文章內容 end-->		
				......</xsl:if>&lt;<a><xsl:attribute name="href">ct.asp?xItem=<xsl:value-of select="@iCuItem" />&amp;ctNode=<xsl:value-of select="../@xNode" />&amp;mp=<xsl:value-of select="//hpMain/mp"/></xsl:attribute>
        		詳全文</a>&gt;
			</p>		
	</xsl:template>
	<!-- 文章內容  此為style3的樣式 end -->

</xsl:stylesheet>






