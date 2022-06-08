<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
	<xsl:include href="headpart.xsl" />
	<xsl:include href="hySearch.xsl" />
	<xsl:include href="epaper.xsl"/>
	<xsl:include href="footer.xsl"/>
	<xsl:include href="menuitem.xsl"/>
	<xsl:include href="xpath.xsl"/>
	<xsl:include href="AttachmentList.xsl" />
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
									<div class="accesskey">
										<a href="accesskey.htm" title="左方區塊" accesskey="L">:::</a>
									</div>
									<!-- accesskey start -->


									<!-- 主要選單 start -->
									<div class="menu">
										<ul>
											<xsl:for-each select="MenuBar">
												<xsl:apply-templates select="MenuCat" />
											</xsl:for-each>
										</ul>
									</div>
									<!-- 主要選單 end -->

								</div>
							</td>
							<!-- leftcol end -->

							<!-- content start -->
							<td id="center">
								<!-- accesskey start -->
								<div class="accesskey">
									<a href="accesskey.htm" title="中央區塊" accesskey="C">:::</a>
								</div>
								<!-- accesskey start -->

								<!-- friendly Start -->
								<div class="path">
									<p>
										<xsl:apply-templates select="xPath" />
									</p>
								</div>
								<div class="friendly">

									<xsl:if test="//commendword/isopen='Y'">
										<div class="back">
											<a>
												<xsl:attribute name="href">
													javascript:getSelectedText('<xsl:value-of select="//qStr" />')
												</xsl:attribute>
												推薦詞彙
											</a>
										</div>
									</xsl:if>


									<div class="print">
										<a>
											<xsl:attribute name="href">
												fp.asp<xsl:value-of select="qStr" />
											</xsl:attribute>
											<xsl:attribute name="target">_blank</xsl:attribute>
											友善列印
										</a>
									</div>
									<div class="email">
										<a>
											<xsl:attribute name="href">
												sp.asp<xsl:value-of select="qStr" />&amp;xdUrl=forward/forward_add.asp
											</xsl:attribute>
											<xsl:attribute name="target">_nwGip</xsl:attribute>
											轉寄友人
										</a>
									</div>
									<div class="email">
										<a>
											<xsl:attribute name="href">
												question_login.asp?mp=<xsl:value-of select="//mp" />|<xsl:value-of select="//MainArticle/@iCuItem"/>
											</xsl:attribute>
											<xsl:attribute name="target">_blank</xsl:attribute>
											我要發問
										</a>
										<div class="email">
											<a href="#" onclick="var domainName=document.domain;window.showModalDialog('http://'+domainName+'/mailbox.asp',self);return false;">
												系統問題
											</a>
											<input type="hidden"  id="type" name="type" value="4"  />
											<input type="hidden" id="ARTICLE_ID"  name="ARTICLE_ID" >
												<xsl:attribute name="value">
													<xsl:value-of select="//MainArticle/@iCuItem"/>
												</xsl:attribute>
											</input>
										</div>
									</div>
                  <DIV align="right">
                    <SPAN class="t12">
                      <SPAN class="gray">調整字級：</SPAN>
                    </SPAN>
                    <A onclick="javascript:changeFontSize('article','idx1');">
                      <IMG id="idx1" alt="" src="./images/fontsize_1_off.gif" align="absMiddle" border="0" name="idx1" />
                    </A>
                    <A onclick="javascript:changeFontSize('article','idx2');">
                      <IMG id="idx2" alt="" src="./images/fontsize_2_off.gif" align="absMiddle" border="0" name="idx2" />
                    </A>
                    <A onclick="javascript:changeFontSize('article','idx3');">
                      <IMG id="idx3" alt="" src="./images/fontsize_3_off.gif" align="absMiddle" border="0" name="idx3" />
                    </A>
                    <A onclick="javascript:changeFontSize('article','idx4');">
                      <IMG id="idx4" alt="" src="./images/fontsize_4_off.gif" align="absMiddle" border="0" name="idx4" />
                    </A>
                  </DIV>
								</div>

                <script type="text/javascript" src="./js/dozoom.js"></script>
								<script type="text/javascript">
									<xsl:text disable-output-escaping="yes">
	<![CDATA[
		function trim(stringToTrim){ return stringToTrim.replace(/^\s+|\s+$/g,"");}
		function getSelectedText(path) {  
			var alertStr = "";
			if (window.getSelection) {         
				// This technique is the most likely to be standardized.         
				// getSelection() returns a Selection object, which we do not document.         
				alertStr = window.getSelection().toString();
				//textarea的處理
				if( alertStr == '' ){
					alertStr = getTextAreaSelection();
				}  				
			}          
			else if (document.getSelection) {         
				// This is an older, simpler technique that returns a string         
				alertStr = document.getSelection();     
			}     
			else if (document.selection) {         
				// This is the IE-specific technique.         
				// We do not document the IE selection property or TextRange objects.         
				alertStr = document.selection.createRange().text;     
			}
			if ( alertStr.length > 10 ) {
				alert("詞彙長度限制10字以內");
			}			
			else {
				alertStr = trim(alertStr);
				window.open(encodeURI("/CommendWord/CommendWordAdd.aspx?type=2&word=" + alertStr + "&" + path),'建議小百科詞彙','resizable=yes,width=565,height=360');
			}			
		} 

		function getTextAreaSelection(){
			var alertStr = '';
			var elementObj = document.getElementsByTagName("textarea");
			var all_length = elementObj.length;      
			for(var i=0 ; i<all_length ; i++){
				if (elementObj[i].selectionStart != undefined && elementObj[i].selectionEnd != undefined) {         
          var start = elementObj[i].selectionStart;         
          var end = elementObj[i].selectionEnd;         
          alertStr = elementObj[i].value.substring(start, end) ;
          elementObj[i].selectionStart = start;
          elementObj[i].selectionEnd = end;
          //將focus指向該element
          elementObj[i].focus();             
				}     
				else alertStr = ''; // Not supported on this browser                                      
			}
			return alertStr ;
		}
	]]>	
	</xsl:text>
								</script>

								<!-- 中央主資料區 Start-->
                                <xsl:apply-templates select="MainArticle" />
                                <xsl:if test="count(AttachmentList/Attachment) != 0">
                                    <xsl:apply-templates select="AttachmentList" mode="listx"/>
                                </xsl:if>
                                <xsl:if test="count(VideoAttachmentList/Attachment) != 0">
                                    <xsl:for-each select="VideoAttachmentList">
                                        <h4>相關影片</h4>
                                        <xsl:for-each select="Attachment">
                                            <embed width="480" height="360" type="application/x-mplayer2" wmode="transparent" showstatusbar="1" autostart="false">
                                                <xsl:attribute name="src">
                                                    <xsl:value-of select="URL" />
                                                </xsl:attribute>
                                            </embed>
                                            <p>
                                                <xsl:value-of select="Caption" />
                                            </p>
                                        </xsl:for-each>
                                    </xsl:for-each>
                                </xsl:if>                
								<!--xsl:apply-templates select="AttachmentList" /-->
								<!-- 中央主資料區 End-->
								<xsl:apply-templates select="pHTML" />
								<div style="text-align: right">
									<a href="#top">
										<img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" border="0"/>
									</a>
								</div>
							</td>
							<!-- content end -->
						</tr>
					</table>
				</div>

				<!-- 頁尾資訊 Start-->
				<xsl:apply-templates select="." mode = "Copyright"/>
				<!-- 頁尾資訊 End-->
			</body>
		</html>
	</xsl:template>




	<!-- 中央主資料區 Start-->
	<xsl:template match="MainArticle">
		<!--  <div class="News"> -->
		<h4>
			<xsl:value-of select="Caption" />
		</h4>
		<xsl:if test="@IsPostDate ='Y'">
			<p>
				<em>
					日期:&amp;nbsp;<xsl:value-of select="PostDate" />
				</em>
			</p>
		</xsl:if>

		<div class="photobox">
			<xsl:if test="@IsPic ='Y'">
				<xsl:if test="xImgFile">
					<img class="leftimg">
						<xsl:attribute name="src">
							<xsl:value-of select="xImgFile" />
						</xsl:attribute>
						<xsl:attribute name="alt">
							<xsl:value-of select="Caption" />
						</xsl:attribute>
					</img>
				</xsl:if>
			</xsl:if>

            <SPAN class="idx2" id="article" name="article">
                <xsl:value-of disable-output-escaping="yes" select="Content" />
            </SPAN>
			<!--<div><a href="#" class="previous">上一頁</a> | <a href="#" class="next">下一頁</a></div> 		-->
		</div>
		<!--</div>-->
	</xsl:template>
	<!-- 中央主資料區 End-->
	<xsl:template match="pHTML">
		<xsl:copy-of select="." />
	</xsl:template>


</xsl:stylesheet>
