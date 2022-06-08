<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
  xmlns:user="urn:user-namespace-here" version="1.0">

	<xsl:template match="hpMain" mode="pageissue"></xsl:template>

	<!-- 原有版型 -->
	<xsl:template match="MainArticle">
		<div id="cp">
			<div class="article">
				<h4>
					<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
				</h4>
				<xsl:if test="MainArticleField[fieldName='ximgFile']">
					<xsl:if test="MainArticleField[fieldName='ximgFile']/Value != ''">
						<div class="phoimg">
							<img>
								<xsl:attribute name="src">
									public/Data/<xsl:value-of select="MainArticleField[fieldName='ximgFile']/Value" />
								</xsl:attribute>
								<xsl:attribute name="alt">
									<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
								</xsl:attribute>
							</img>
						</div>
					</xsl:if>
				</xsl:if>
				<p>
				<span id="article" name="article" class="idx1">
					<xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xbody']/Value" />
				</span>
				</p>
				<!--h5><xsl:value-of select="MainArticleField[fieldName='stitle']/Value" /></h5>
				<br/>
				<p>
				<em>圖說:</em><br/>
				<xsl:value-of select="MainArticleField[fieldName='xbody']/Value" /></p-->
			</div>
		</div>
		<xsl:apply-templates select="//AttachmentList" />
	</xsl:template>

	<xsl:template match="AttachmentList">
		<div class="lpphoto2" >
			<h5>相關照片</h5>
				<div class="yoxview">
					<ul>
						<xsl:for-each select="Attachment">
							<xsl:apply-templates select="." />
						</xsl:for-each>
					</ul>
				</div>
			</div>
	</xsl:template>

	<xsl:template match="Attachment">
		<li>
			<xsl:choose>
				<xsl:when test="IsImageFile='Y'">
					<a>
						<xsl:attribute name="href">
							<xsl:value-of select="URL" />
						</xsl:attribute>
						<img>
							<xsl:attribute name="src">
								<xsl:value-of select="URL" />
							</xsl:attribute>
							<xsl:attribute name="alt">
								<xsl:value-of select="Caption" />
							</xsl:attribute>
							<xsl:attribute name="title">
								<xsl:value-of select="Caption" />
							</xsl:attribute>
						</img>
						<div style="clear: both;" />
						<div align="center" >
							<xsl:choose>
								<xsl:when test="Descxx!=''">
									<p>
										<xsl:value-of select="Descxx" />
									</p>
								</xsl:when>
								<xsl:otherwise>
									<p>
										<xsl:value-of select="Caption" />
									</p>
								</xsl:otherwise>
							</xsl:choose>
						</div>
					</a>
				</xsl:when>
				<xsl:otherwise>
					<a target="_nwGIP">
						<xsl:attribute name="href">
							<xsl:value-of select="URL" />
						</xsl:attribute>
						<xsl:attribute name="title">
							<xsl:value-of select="Caption" />
						</xsl:attribute>
						<img src="/public/file.png" style="height:96px;width:96px" border="0" />
					</a>
					<div style="clear: both;" />
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
				</xsl:otherwise>
			</xsl:choose>
		</li>
	</xsl:template>
	<!-- 原有版型結束 -->

	<!-- 五合一版型之繼承版型-->
	<xsl:template match="MainArticle" mode="styleA">
		<div id="cp">
			<div class="article">
				<h4>
					<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
				</h4>
				<xsl:if test="MainArticleField[fieldName='ximgFile']/Value!=''">
					<div class="phoimg">
						<img>
							<xsl:attribute name="src">
								public/Data/<xsl:value-of select="MainArticleField[fieldName='ximgFile']/Value" />
							</xsl:attribute>
							<xsl:attribute name="alt">
								<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
							</xsl:attribute>
						</img>
					</div>
				</xsl:if>
				<p><span id="article" name="article" class="idx1">
					<xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xbody']/Value" />
					</span>
				</p>
			</div>
		</div>
		<xsl:apply-templates select="//AttachmentList" mode="styleA" />
	</xsl:template>
	<xsl:template match="AttachmentList" mode="styleA">
	<div class="lpphoto2">
		<h5>相關照片</h5>
			<div id="contentPanel">
                 <div id="overviewLeft" class="left">
                     <div class="yoxview">
						<ul>
							<xsl:for-each select="Attachment">
								<xsl:apply-templates select="." mode="styleA" />
							</xsl:for-each>
						</ul>
					</div>
				</div>
			</div>
		</div>
	</xsl:template>
	<xsl:template match="Attachment" mode="styleA">
		<li>
			<xsl:choose>
				<xsl:when test="IsImageFile='Y'">
					<a>
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
							<xsl:attribute name="title">
								<xsl:value-of select="Caption" />
							</xsl:attribute>					
						</img>
						<div style="clear: both;" />
						<div align="center" >
							<xsl:choose>
								<xsl:when test="Descxx!=''">
									<p>
										<xsl:value-of select="Descxx" />
									</p>
								</xsl:when>
								<xsl:otherwise>
									<p>
										<xsl:value-of select="Caption" />
									</p>
								</xsl:otherwise>
							</xsl:choose>
						</div>
					</a>	
				</xsl:when>
				<xsl:otherwise>
					<a target="_nwGIP">
						<xsl:attribute name="href">
							<xsl:value-of select="URL" />
						</xsl:attribute>
						<xsl:attribute name="title">
							<xsl:value-of select="Caption" />
						</xsl:attribute>
						<img src="/public/file.png" style="height:96px;width:96px" border="0" />
					</a>
					<div style="clear: both;" />
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
				</xsl:otherwise>
			</xsl:choose>
		</li>
	</xsl:template>
	<!-- 五合一版型之繼承版型結束 -->

	<!-- 五合一版型之第一個特殊版 -->
	<xsl:template match="MainArticle" mode="styleB">
		<div class="cp08">
			<div class="template1">
				<h4>
					<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
				</h4>
				<div class="column">
					<xsl:if test="MainArticleField[fieldName='ximgFile']/Value!=''">
						 <div class="phoimg yoxview">
						 <a>
						 <xsl:attribute name="href">
							/public/Data/<xsl:value-of select="MainArticleField[fieldName='ximgFile']/Value" />
						</xsl:attribute>
						<xsl:attribute name="title">
							<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
						</xsl:attribute>
							<img>
								<xsl:attribute name="src">
									/public/Data/<xsl:value-of select="MainArticleField[fieldName='ximgFile']/Value" />
								</xsl:attribute>
								<xsl:attribute name="alt">
									<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
								</xsl:attribute>
							</img>
						</a>	
						</div>
					</xsl:if>
					<p><span id="article" name="article" class="idx1">
						<xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xbody']/Value" />
						</span>
					</p>
					</div>
						<div class="lpphoto2 yoxview">
							<xsl:apply-templates select="//AttachmentList" mode="styleB" />
						</div>
					</div>
				</div>
	</xsl:template>
	<!-- 五合一版型之第一個特殊版的附件 -->
	<xsl:template match="AttachmentList" mode="styleB">
		<div class="attachment">
			<h5>附件</h5>
			<ul>
				<xsl:for-each select="Attachment">
					<li>
						<xsl:if test="fileType='jpg' or fileType='gif'">
						<a>
							<xsl:attribute name="href">
								<xsl:value-of select="URL" />
							</xsl:attribute>
							<xsl:attribute name="title">
								<xsl:value-of select="Caption" />
							</xsl:attribute>
								<div class="image">
									<img>
										<xsl:attribute name="src">
											<xsl:value-of select="URL" />
										</xsl:attribute>
										<xsl:attribute name="alt">
											<xsl:value-of select="Caption" />
										</xsl:attribute>
										<xsl:attribute name="title">
											<xsl:value-of select="Caption" />
										</xsl:attribute>					
									</img>
								</div>
							<xsl:value-of select="Caption" />
						</a>
							
						</xsl:if>
						<xsl:if test="fileType!='jpg' and fileType!='gif'">
						<div class="title">
							<a target="_nwGIP">
								<xsl:attribute name="href">
									<xsl:value-of select="URL" />
								</xsl:attribute>
								<xsl:value-of select="Caption" />
							</a>
						</div>
						</xsl:if>
						<p>
							<xsl:value-of select="Descxx" />
						</p>
					</li>
				</xsl:for-each>
				<!--End-->
			</ul>
		</div>
	</xsl:template>
	<!-- 五合一版型之第一個特殊版 結束 -->

	<!-- 五合一版型之第二個特殊版 -->
	<xsl:template match="MainArticle" mode="styleC">
		<div class="cp08">
			<div class="template2">
				<h4>
					<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
				</h4>
				<xsl:if test="count(//AttachmentList/Attachment) > 0">
					<div class="column">
						<xsl:if test="MainArticleField[fieldName='ximgFile']/Value!=''">
							<div class="image">
								<img>
									<xsl:attribute name="src">
										/public/Data/<xsl:value-of select="MainArticleField[fieldName='ximgFile']/Value" />
									</xsl:attribute>
									<xsl:attribute name="alt">
										<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
									</xsl:attribute>
								</img>
							</div>
						</xsl:if>
						<p><span id="article" name="article" class="idx1">
							<xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xbody']/Value" />
							</span>
						</p>
					</div>
				</xsl:if>								
				<xsl:if test="count(//AttachmentList/Attachment) = 0">
					<div class="column1">						
						<xsl:if test="MainArticleField[fieldName='ximgFile']/Value!=''">
							<div class="image">
								<img>
									<xsl:attribute name="src">
										/public/Data/<xsl:value-of select="MainArticleField[fieldName='ximgFile']/Value" />
									</xsl:attribute>
									<xsl:attribute name="alt">
										<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
									</xsl:attribute>
								</img>
							</div>
						</xsl:if>
						<p><span id="article" name="article" class="idx1">
							<xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xbody']/Value" />
							</span>
						</p>
					</div>
				</xsl:if>				
				<div class="thumb">
				<div class="yoxview">
					<xsl:for-each select="//AttachmentList/Attachment">
							<xsl:if test="fileType='jpg' or fileType='gif'">
								<ul>
									<xsl:if test="picCount &lt; 6">
										<li>
											<div class="image">
												<a>
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
														<xsl:attribute name="title">
															<xsl:value-of select="Caption" />
														</xsl:attribute>					
													</img>
												</a>
											</div>
										<xsl:choose>
											<xsl:when test="Descxx!=''">
												<p>
													<xsl:value-of select="Descxx" />
												</p>
											</xsl:when>
											<xsl:otherwise>
												<p>
													<xsl:value-of select="Caption" />
												</p>
											</xsl:otherwise>
										</xsl:choose>
									</li>
								</xsl:if>
							</ul>
						</xsl:if>
					</xsl:for-each>
					</div>
				</div>
				<!--End-->
				<!--xsl:apply-templates select="//AttachmentList" mode="styleC" /-->
			</div>
		</div>
	</xsl:template>
	<!-- 五合一版型之第二個特殊版的附件 -->
	<xsl:template match="AttachmentList" mode="styleC">
		<div class="attachment">
			<h5>附件</h5>
			<ul>
				<xsl:for-each select="Attachment">
					<li>
						<div class="title">
							<a target="_nwGIP">
								<xsl:attribute name="href">
									<xsl:value-of select="URL" />
								</xsl:attribute>
								<xsl:value-of select="Caption" />
							</a>
						</div>
					</li>
				</xsl:for-each>
			</ul>
		</div>
	</xsl:template>
	<!-- 五合一版型之第二個特殊版 結束 -->

	<!-- 五合一版型之第三個特殊版 -->
	<xsl:template match="MainArticle" mode="styleD">
		<div class="cp08">
			<div class="template3">
				<h4>
					<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
				</h4>
			<div class="yoxview">
				<div class="column">
					<div class="thumb1">
					
						<ul>
							<xsl:for-each select="//AttachmentList/Attachment">
								<xsl:if test="picCount &lt; 3">
									<li>
										<div class="image">
											<xsl:choose>
												<xsl:when test="IsImageFile='Y'">
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
															<xsl:attribute name="title">
																<xsl:value-of select="Caption" />
															</xsl:attribute>
														</img>
													</a>
												</xsl:when>
												<xsl:otherwise>
													<a target="_nwGIP">
														<xsl:attribute name="href">
															<xsl:value-of select="URL" />
														</xsl:attribute>
														<xsl:attribute name="title">
															<xsl:value-of select="Caption" />
														</xsl:attribute>
														<img src="/public/file.png" style="height:96px;width:96px" border="0" />
													</a>
												</xsl:otherwise>
											</xsl:choose>
										</div>
										<xsl:choose>
											<xsl:when test="Descxx!=''">
												<p>
													<xsl:value-of select="Descxx" />
												</p>
											</xsl:when>
											<xsl:otherwise>
												<p>
													<xsl:value-of select="Caption" />
												</p>
											</xsl:otherwise>
										</xsl:choose>
									</li>
								</xsl:if>
							</xsl:for-each>
						</ul>
					</div>
					<div class="image">
						<xsl:if test="MainArticleField[fieldName='ximgFile']/Value!=''">
							<img>
								<xsl:attribute name="src">/public/Data/<xsl:value-of select="MainArticleField[fieldName='ximgFile']/Value" /></xsl:attribute>
								<xsl:attribute name="alt">
									<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
								</xsl:attribute>
							</img>
						</xsl:if>
					</div>
					<div class="thumb2">
						<ul>
							<xsl:for-each select="//AttachmentList/Attachment">
								<xsl:if test="picCount &lt; 7 and picCount &gt; 2">
									<li>
										<div class="image">
											<xsl:choose>
												<xsl:when test="IsImageFile='Y'">
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
															<xsl:attribute name="title">
																<xsl:value-of select="Caption" />
															</xsl:attribute>
														</img>
													</a>
												</xsl:when>
												<xsl:otherwise>
													<a target="_nwGIP">
														<xsl:attribute name="href">
															<xsl:value-of select="URL" />
														</xsl:attribute>
														<xsl:attribute name="title">
															<xsl:value-of select="Caption" />
														</xsl:attribute>
														<img src="/public/file.png" style="height:96px;width:96px" border="0" />
													</a>
												</xsl:otherwise>
											</xsl:choose>
										</div>
										<xsl:choose>
											<xsl:when test="Descxx!=''">
												<p>
													<xsl:value-of select="Descxx" />
												</p>
											</xsl:when>
											<xsl:otherwise>
												<p>
													<xsl:value-of select="Caption" />
												</p>
											</xsl:otherwise>
										</xsl:choose>
									</li>
								</xsl:if>
							</xsl:for-each>
						</ul>
					</div>
					<p>
					<span id="article" name="article" class="idx1">
						<xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xbody']/Value" />
					</span>
					</p>
				</div>
				</div>
				<xsl:apply-templates select="//AttachmentList" mode="styleD" />
			</div>
		<!--End-->	
		</div>
	</xsl:template>
	<!-- 五合一版型之第三個特殊版的附件 -->
	<xsl:template match="AttachmentList" mode="styleD">
		<div class="attachment">
			<h5>附件</h5>
			<ul>
				<xsl:for-each select="Attachment">
					<xsl:if test="picCount='' or picCount &gt; 6">
						<li>
							<div class="title">
								<a target="_nwGIP">
									<xsl:attribute name="href">
										<xsl:value-of select="URL" />
									</xsl:attribute>
									<xsl:value-of select="Caption" />
								</a>
							</div>
						</li>
					</xsl:if>
				</xsl:for-each>
			</ul>
		</div>
	</xsl:template>
	<!-- 五合一版型之第三個特殊版 結束 -->

	<!-- 五合一版型之第四個特殊版  -->
	<xsl:template match="MainArticle" mode="styleE">
	
		<div class="cp08">
			<div class="template4">
				<h4>
					<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
				</h4>
				<div class="column">
					<div class="image">
						<xsl:if test="MainArticleField[fieldName='ximgFile']/Value!=''">
							<img>
								<xsl:attribute name="src">
									/public/Data/<xsl:value-of select="MainArticleField[fieldName='ximgFile']/Value" />
								</xsl:attribute>
								<xsl:attribute name="alt">
									<xsl:value-of select="MainArticleField[fieldName='stitle']/Value" />
								</xsl:attribute>
							</img>
						</xsl:if>
					</div>
					<p>
					<span id="article" name="article" class="idx1">
						<xsl:value-of disable-output-escaping="yes" select="MainArticleField[fieldName='xbody']/Value" />
					</span>
					</p>
				</div>
				<xsl:apply-templates select="//AttachmentList" mode="styleE" />
			</div>
		</div>
	</xsl:template>
	<!-- 五合一版型之第四個特殊版的附件 -->
	<xsl:template match="AttachmentList" mode="styleE">
		<div class="attachment">
			<h4>附件</h4>
			<xsl:if test="contains(allFileType,'wmv') or contains(allFileType,'asf') or contains(allFileType,'wma') or contains(allFileType,'mp3')">
				<div class="type1">
					<h5>影音檔</h5>
					<ul>
						<xsl:for-each select="Attachment">
							<xsl:if test="fileType='wmv' or fileType='asf' or fileType='wma' or fileType='mp3'">
								<li>
									<a target="_nwGIP">
										<xsl:attribute name="href">
											<xsl:value-of select="URL" />
										</xsl:attribute>
										<xsl:value-of select="Caption" />
									</a>
								</li>
							</xsl:if>
						</xsl:for-each>
					</ul>
				</div>
			</xsl:if>
			<xsl:if test="contains(allFileType,'jpg') or contains(allFileType,'gif')">
				<div class="type2">
					<h5>圖片檔</h5>

						<ul>
							<xsl:for-each select="Attachment">
								<xsl:if test="fileType='jpg' or fileType='gif'">
									<li>
										<a target="_nwGIP">
											<xsl:attribute name="href">
												<xsl:value-of select="URL" />
											</xsl:attribute>
											
										<xsl:value-of select="Caption" />
									</a>
								</li>
							</xsl:if>
						</xsl:for-each>
					</ul>
				</div>
			</xsl:if>
			<xsl:if test="contains(allFileType,'doc') or contains(allFileType,'pdf')">
				<div class="type2">
					<h5>文件檔</h5>
					<ul>
						<xsl:for-each select="Attachment">
							<xsl:if test="fileType='doc' or fileType='pdf'">
								<li>
									<a target="_nwGIP">
										<xsl:attribute name="href">
											<xsl:value-of select="URL" />
										</xsl:attribute>
										<xsl:value-of select="Caption" />
									</a>
								</li>
							</xsl:if>
						</xsl:for-each>
					</ul>
				</div>
			</xsl:if>
		</div>
	</xsl:template>
	<!-- 五合一版型之第四個特殊版 結束 -->

</xsl:stylesheet>

