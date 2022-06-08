<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
    xmlns:user="urn:user-namespace-here" version="1.0"> <!-- 網站標頭／開始 -->
    <xsl:template match="hpMain" mode="topinfo">
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>
            <xsl:value-of select="$title" />
        </title>
    </xsl:template> <!-- 網站標頭／結束 --> <!-- 導覽列／開始 -->
    <xsl:template match="hpMain" mode="submenu">
        <xsl:for-each select="//BlockF/Article">
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
            </li>
        </xsl:for-each>
    </xsl:template> <!-- 導覽列／結束 --> <!-- 連絡我們／開始 -->
    <xsl:template match="hpMain" mode="ContactUs">
        <xsl:attribute name="href">form.asp?ctUnit=113&amp;spec=contactus</xsl:attribute>
        <xsl:attribute name="target">_nwGip</xsl:attribute>
    </xsl:template> <!-- 連絡我們／結束 --> <!--  路徑連結／開始 -->
    <xsl:template match="hpMain" mode="xsPath">
        <div class="Path">
            瀏覽路徑：<a>
            <xsl:attribute name="href">dp.asp?mp=<xsl:value-of select="//mp" /></xsl:attribute>
            
            回首頁</a>
            <xsl:for-each select="//xPath/xPathNode">
                <img>
                    <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/path_arrow.gif</xsl:attribute>
                    <xsl:attribute name="alt">
                        <xsl:value-of select="//xxarrow" />
                    </xsl:attribute>
                </img>
                <a>
                    <xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="@xNode" />&amp;xp1=<xsl:value-of select="//Mp" /></xsl:attribute>
                    <xsl:attribute name="title">
                        <xsl:value-of select="@Title" />
                    </xsl:attribute>
                    <xsl:value-of select="@Title" />
                </a>
            </xsl:for-each>
        </div>
    </xsl:template> <!--  路徑連結／結束 --> 
    <!-- 資料大類／開始 -->
    <xsl:template match="hpMain" mode="xsCatList">
        <xsl:if test="CatList">
            <ul class="Category">
                <li>
                    <a>
                        <xsl:attribute name="href">lp.asp?<xsl:value-of select="//xqURL" /></xsl:attribute>
                        <xsl:attribute name="title">All</xsl:attribute>
                        All
                    </a>
                </li>
                <xsl:for-each select="//CatList/CatItem">
                    <li>
                        <a>
                            <xsl:attribute name="href">lp.asp?<xsl:value-of select="//xqURL" /><xsl:value-of select="xqCondition" /></xsl:attribute>
                            <xsl:attribute name="title">
                                <xsl:value-of select="CatName" />
                            </xsl:attribute>
                            <xsl:value-of select="CatName" />
                        </a>
                    </li>
                </xsl:for-each>
            </ul>
        </xsl:if>
    </xsl:template>
    <!-- 資料大類／結束 -->
    <!-- 頁尾資訊／開始 -->
    <xsl:template match="hpMain" mode="bottominfo">
        <br />
			<img src="xslgip/images_data/60x80_good.gif"/>
        <p>行政院農業委員會 版權所有c 2004 COA All Rights Reserved‧<br/>
        維護單位：<a href="http://www.caes.gov.tw/">農業試驗所(嘉義分所)</a><br/>
        最佳瀏覽狀態為 IE4.0 以上‧1024*768 解析度最佳</p>
    </xsl:template>
    <!-- 頁尾資訊／結束 -->
    <!-- 選單／開始 -->
    <xsl:template match="hpMain" mode="xsMenu">
        <ul class="MainMenu">
            <xsl:for-each select="MenuBar/MenuCat">
                <li>
                    <a>
                        <xsl:attribute name="href">
                            <xsl:value-of select="redirectURL" />
                        </xsl:attribute>
                        <xsl:if test="@newWindow='Y'">
                            <xsl:attribute name="target">_nwGip</xsl:attribute>
                        </xsl:if>
                        <xsl:attribute name="title">
                            <xsl:value-of select="Caption" />
                        </xsl:attribute>
                        <xsl:value-of select="Caption" />
                    </a>
                </li>
                <xsl:if test="MenuItem">
                    <ul>
                        <xsl:for-each select="MenuItem">
                            <xsl:apply-templates select="." />
                        </xsl:for-each>
                    </ul>
                </xsl:if> <!--
                <xsl:if test="MenuItem1">
				<ul>
					<xsl:for-each select="MenuItem1">
						<xsl:apply-templates select="." />				
					</xsl:for-each>
				</ul>
			</xsl:if>
			-->
            </xsl:for-each>
        </ul>
    </xsl:template>
    <xsl:template match="hpMain" mode="Menu2">
        <ul class="Menu">
            <xsl:choose>
                <xsl:when test="MenuBar/MenuCat/MenuItem1">
                    <xsl:for-each select="MenuBar/MenuCat/MenuItem1">
                        <xsl:apply-templates select="." />
                    </xsl:for-each>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:for-each select="MenuBar/MenuCat">
                        <li>
                            <a>
                                <xsl:attribute name="href">
                                    <xsl:value-of select="redirectURL" />
                                </xsl:attribute>
                                <xsl:if test="@newWindow='Y'">
                                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                                </xsl:if>
                                <xsl:attribute name="title">
                                    <xsl:value-of select="Caption" />
                                </xsl:attribute>
                                <xsl:value-of select="Caption" />
                            </a>
                            <xsl:if test="MenuItem">
                                <ul>
                                    <xsl:for-each select="MenuItem">
                                        <xsl:apply-templates select="." />
                                    </xsl:for-each>
                                </ul>
                            </xsl:if>
                        </li>
                    </xsl:for-each>
                </xsl:otherwise>
            </xsl:choose>
        </ul>
<div class="Menu_Foot" xmlns=""></div>
    </xsl:template>
    <xsl:template match="hpMain" mode="Menu">
        <ul class="Menu">
            <xsl:for-each select="MenuBar/MenuCat">
                <li>
                    <a>
                        <xsl:attribute name="href">
                            <xsl:value-of select="redirectURL" />
                        </xsl:attribute>
                        <xsl:if test="@newWindow='Y'">
                            <xsl:attribute name="target">_nwGip</xsl:attribute>
                        </xsl:if>
                        <xsl:attribute name="title">
                            <xsl:value-of select="Caption" />
                        </xsl:attribute>
                        <xsl:value-of select="Caption" />
                    </a>
                    <xsl:if test="MenuItem">
                        <ul>
                            <xsl:for-each select="MenuItem">
                                <xsl:apply-templates select="." />
                            </xsl:for-each>
                        </ul>
                    </xsl:if>
                </li> <!--
                <xsl:if test="MenuItem1">
				    <ul>
					    <xsl:for-each select="MenuItem1">
						    <xsl:apply-templates select="." />				
					    </xsl:for-each>
				    </ul>
			    </xsl:if>
			    -->
            </xsl:for-each>
        </ul>
	<div class="Menu_Foot" xmlns=""></div>
    </xsl:template>
    <xsl:template match="MenuItem">
        <li>
            <a>
                <xsl:attribute name="href">
                    <xsl:value-of select="redirectURL" />
                </xsl:attribute>
                <xsl:if test="@newWindow='Y'">
                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                </xsl:if>
                <xsl:attribute name="title">
                    <xsl:value-of select="Caption" />
                </xsl:attribute>
                <xsl:value-of select="Caption" />
            </a>
        </li>
    </xsl:template>
    <xsl:template match="MenuItem1">
        <li>
            <a>
                <xsl:attribute name="href">
                    <xsl:value-of select="redirectURL" />
                </xsl:attribute>
                <xsl:if test="@newWindow='Y'">
                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                </xsl:if>
                <xsl:attribute name="title">
                    <xsl:value-of select="Caption" />
                </xsl:attribute>
                <xsl:value-of select="Caption" />
            </a>
        </li>
        <xsl:if test="MenuItem2">
            <ul>
                <xsl:for-each select="MenuItem2">
                    <xsl:apply-templates select="." />
                </xsl:for-each>
            </ul>
        </xsl:if>
    </xsl:template>
    <xsl:template match="MenuItem2">
        <li>
            <a>
                <xsl:attribute name="href">
                    <xsl:value-of select="redirectURL" />
                </xsl:attribute>
                <xsl:if test="@newWindow='Y'">
                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                </xsl:if>
                <xsl:attribute name="title">
                    <xsl:value-of select="Caption" />
                </xsl:attribute>
                <xsl:value-of select="Caption" />
            </a>
        </li>
    </xsl:template> <!-- 選單／結束 --> <!-- RSS，查詢本單元／開始 -->
    <xsl:template match="hpMain" mode="xsFunctions">
        <xsl:if test="//YNrss='Y' or //YNquery='Y'">
            <ul class="Function2">
                <xsl:if test="//YNrss='Y'">
                    <li>
                        <a class="RSS"> <!--
						<xsl:attribute name="href">rss.asp?mp=<xsl:value-of select="//Mp" /></xsl:attribute>
						 -->
                            <xsl:attribute name="href">ws_RSS.asp?ctNodeId=<xsl:value-of select="//TopicList/@xNode" /></xsl:attribute>
                            <xsl:attribute name="title">
                                <xsl:value-of select="//xxrss" />
                            </xsl:attribute>
                            <xsl:value-of select="//xxrss" />
                        </a>
                    </li>
                    <link rel="alternate" type="application/rss+xml" title="">
                        <xsl:attribute name="href">http://ugip.hyweb.com.tw/site/giotw/ws_RSS.asp?ctNodeId=<xsl:value-of select="//TopicList/@xNode" /></xsl:attribute>
                        <xsl:attribute name="title">
                            <xsl:value-of select="//xPath/UnitName" />
                        </xsl:attribute>
                    </link>
                </xsl:if>
                <xsl:if test="//YNquery='Y'">
                    <li>
                        <a class="GIPsearch">
                            <xsl:attribute name="href">qp.asp?<xsl:value-of select="xqURL" /></xsl:attribute>
                            <xsl:attribute name="title">
                                <xsl:value-of select="//xxquery" />
                            </xsl:attribute>
                            <xsl:value-of select="//xxquery" />
                        </a>
                    </li>
                </xsl:if>
            </ul>
        </xsl:if>
    </xsl:template>
    <!-- RSS，查詢本單元／結束 -->
    <!-- 其他功能／開始 -->
    <xsl:template match="hpMain" mode="Functions">
        <ul class="Function">

<xsl:if test="//commendword/isopen='Y'">
<li>
<a class="Forward">
<xsl:attribute name="href">javascript:getSelectedText('<xsl:value-of select="//qStr" />')</xsl:attribute>
推薦詞彙
</a>
</li>
</xsl:if>


            <li>
                <a class="Print">
                    <xsl:attribute name="href">fp.asp?xItem=<xsl:value-of select="//@iCuItem" />&amp;ctNode=<xsl:value-of select="//xPathNode[position()=last()]/@xNode" /></xsl:attribute>
                    <!--xsl:attribute name="href">#</xsl:attribute-->
                    <xsl:attribute name="title">友善列印</xsl:attribute>
                    <xsl:attribute name="target">_blank</xsl:attribute>
                    友善列印
                </a>
            </li>
            <!--<li>
                <a class="Forward">
                    <xsl:attribute name="href">form.asp?xItem=<xsl:value-of select="//@iCuItem" />&amp;ctNode=<xsl:value-of select="//xPathNode[position()=last()]/@xNode" />&amp;spec=forward</xsl:attribute>
                    <xsl:attribute name="href">#</xsl:attribute>
                    <xsl:attribute name="title">轉寄文章</xsl:attribute>
                    <xsl:attribute name="target">_blank</xsl:attribute>
                    轉寄文章
                </a>
            </li>-->
            <li>
                <a href="javascript:history.go(-1);" class="Back">
                    <xsl:attribute name="title">回上一頁</xsl:attribute>
                    回上一頁
                </a>
                <noscript>本網頁使用SCRIPT編碼方式執行回上一頁的動作，如果您的瀏覽器不支援SCRIPT，請直接使用「Alt+方向鍵向左按鍵」來返回上一頁</noscript>
            </li>
			<li>
				<a class="Forward">
                    <xsl:attribute name="href">question_login.asp?mp=<xsl:value-of select="//mp" />|<xsl:value-of select="//MainArticle/@iCuItem"/></xsl:attribute>
                    <xsl:attribute name="title">我要發問</xsl:attribute>
                    <xsl:attribute name="target">_blank</xsl:attribute>
                    我要發問
                </a>
			</li>
			<li>
				<a class="Forward" href="#" onclick="var domainName=document.domain;window.showModalDialog('http://'+domainName+'/mailbox.asp',self);return false;">
												系統問題
				</a>
				<input type="hidden"  name="type" value="4"  />
				<input type="hidden" name="ARTICLE_ID" >
				<xsl:attribute name="value">
				<xsl:value-of select="//MainArticle/@iCuItem"/>
				</xsl:attribute>
				</input>
			</li>
        </ul>

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


    </xsl:template>
    <!-- 其他功能／結束 -->
    <!-- 廣告／開始-->
    <xsl:template match="hpMain" mode="xsAD">
        <div class="AD">
            <xsl:for-each select="//AD/Article">
                <a>
                    <xsl:attribute name="href">
                        <xsl:value-of select="xURL" />
                    </xsl:attribute>
                    <xsl:if test="@newWindow='Y'">
                        <xsl:attribute name="target">_nwGip</xsl:attribute>
                    </xsl:if>
                    <xsl:if test="xImgFile">x
                        <img border="0">
                            <xsl:attribute name="src"><xsl:value-of select="xImgFile" />
                            </xsl:attribute>
                            <xsl:attribute name="alt">
                                <xsl:value-of select="Caption" />
                            </xsl:attribute>
                        </img>
                    </xsl:if>
                </a>
            </xsl:for-each>
        </div>
    </xsl:template>
    <!-- 廣告／結束-->
    <!--  網站地圖路徑連結／開始  -->
    <xsl:template match="hpMain" mode="SitemapxPath">
        <div class="Path">
            <xsl:value-of select="//xxpath" />
            <a>
                <xsl:attribute name="href">dp.asp?mp=<xsl:value-of select="//Mp" /></xsl:attribute>
                <xsl:attribute name="title">
                    <xsl:value-of select="//xxhome" />
                </xsl:attribute>
                <xsl:value-of select="//xxhome" />
            </a>
            <img>
                <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/path_arrow.gif</xsl:attribute>
                <xsl:attribute name="alt">
                    <xsl:value-of select="//xxarrow" />
                </xsl:attribute>
            </img>
            <xsl:value-of select="SitemapTop/Caption" />
        </div>
    </xsl:template>
    <!-- 網站地圖 路徑連結／結束  -->
    <!--  網站地圖資訊／開始  -->
    <xsl:template match="hpMain" mode="SitemapTop">
        <h3>
            <xsl:value-of select="SitemapTop/Article /Caption" />
        </h3>
        <xsl:value-of disable-output-escaping="yes" select="SitemapTop/Article/ContentAll" />
    </xsl:template>
    <!--  網站地圖資訊／結束  -->
    <!-- 網站導覽 -->
    <xsl:template match="hpMain" mode="Sitemap">
        <h3>
            <xsl:value-of select="SitemapTop/Caption" /> <!--xsl:value-of select="//CtUnitName"/--></h3>
        <ol> <!--第一層 -->
            <xsl:for-each select="Sitemap/Menucat[@dataparent='0']">
                <xsl:variable name="thisPosition">
                    <xsl:value-of select="position()" />
                </xsl:variable>
                <li>
                    <xsl:choose>
                        <xsl:when test="xURL!=''">
						<xsl:value-of select="$thisPosition" />.
						<a>
                                <xsl:attribute name="href">
                                    <xsl:value-of select="xURL" />
                                </xsl:attribute>
                                <xsl:attribute name="title">
                                    <xsl:value-of select="Caption" />
                                </xsl:attribute>
                                <xsl:if test="xURL/@newWindow='Y'">
                                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                                </xsl:if>
                                <xsl:value-of select="Caption" />
                            </a>
					</xsl:when>
                        <xsl:otherwise>
						<xsl:value-of select="$thisPosition" />.
						<xsl:value-of select="Caption" />
					</xsl:otherwise>
                    </xsl:choose> <!--第二層 -->
                    <xsl:variable name="thisCaptionID">
                        <xsl:value-of select="@iCuItem" />
                    </xsl:variable>
                    <xsl:if test="@isUnit='C'">
                        <ol>
                            <xsl:for-each select="ancestor::hpMain//Sitemap/Menucat[@dataparent=$thisCaptionID]">
                                <li>
                                    <xsl:variable name="thisPosition2">
                                        <xsl:value-of select="position()" />
                                    </xsl:variable>
                                    <xsl:choose>
                                        <xsl:when test="xURL!=''">
									<xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> .
									<a>
                                                <xsl:attribute name="href">
                                                    <xsl:value-of select="xURL" />
                                                </xsl:attribute>
                                                <xsl:attribute name="title">
                                                    <xsl:value-of select="Caption" />
                                                </xsl:attribute>
                                                <xsl:if test="xURL/@newWindow='Y'">
                                                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                                                </xsl:if>
                                                <xsl:value-of select="Caption" />
                                            </a>
								</xsl:when>
                                        <xsl:otherwise>
									<xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> .
									<xsl:value-of select="Caption" />
								</xsl:otherwise>
                                    </xsl:choose> <!--第三層 -->
                                    <xsl:variable name="thisCaptionID2">
                                        <xsl:value-of select="@iCuItem" />
                                    </xsl:variable>
                                    <xsl:if test="@isUnit='C'">
                                        <ol>
                                            <xsl:for-each select="ancestor::hpMain//Sitemap/Menucat[@dataparent=$thisCaptionID2]">
                                                <li>
                                                    <xsl:variable name="thisPosition3">
                                                        <xsl:value-of select="position()" />
                                                    </xsl:variable>
                                                    <xsl:choose>
                                                        <xsl:when test="xURL!=''">
												<xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> - <xsl:value-of select="$thisPosition3" /> .
												<a>
                                                                <xsl:attribute name="href">
                                                                    <xsl:value-of select="xURL" />
                                                                </xsl:attribute>
                                                                <xsl:attribute name="title">
                                                                    <xsl:value-of select="Caption" />
                                                                </xsl:attribute>
                                                                <xsl:if test="xURL/@newWindow='Y'">
                                                                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                                                                </xsl:if>
                                                                <xsl:value-of select="Caption" />
                                                            </a>
											</xsl:when>
                                                        <xsl:otherwise>
												<xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> - <xsl:value-of select="$thisPosition3" /> .
												<xsl:value-of select="Caption" />
											</xsl:otherwise>
                                                    </xsl:choose> <!--第四層 -->
                                                    <xsl:variable name="thisCaptionID3">
                                                        <xsl:value-of select="@iCuItem" />
                                                    </xsl:variable>
                                                    <xsl:if test="@isUnit='C'">
                                                        <ol>
                                                            <xsl:for-each select="ancestor::hpMain//Sitemap/Menucat[@dataparent=$thisCaptionID3]">
                                                                <li>
                                                                    <xsl:variable name="thisPosition4">
                                                                        <xsl:value-of select="position()" />
                                                                    </xsl:variable>
                                                                    <xsl:choose>
                                                                        <xsl:when test="xURL!=''">
															<xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> -
															<xsl:value-of select="$thisPosition3" /> - <xsl:value-of select="$thisPosition4" /> .
															<a>
                                                                                <xsl:attribute name="href">
                                                                                    <xsl:value-of select="xURL" />
                                                                                </xsl:attribute>
                                                                                <xsl:attribute name="title">
                                                                                    <xsl:value-of select="Caption" />
                                                                                </xsl:attribute>
                                                                                <xsl:if test="xURL/@newWindow='Y'">
                                                                                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                                                                                </xsl:if>
                                                                                <xsl:value-of select="Caption" />
                                                                            </a>
														</xsl:when>
                                                                        <xsl:otherwise>
															<xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> -
															<xsl:value-of select="$thisPosition3" /> - <xsl:value-of select="$thisPosition4" /> .
															<xsl:value-of select="Caption" />
														</xsl:otherwise>
                                                                    </xsl:choose>
                                                                </li>
                                                            </xsl:for-each>
                                                        </ol>
                                                    </xsl:if> <!--第四層 -->
                                                </li>
                                            </xsl:for-each>
                                        </ol>
                                    </xsl:if> <!--第三層 -->
                                </li>
                            </xsl:for-each>
                        </ol>
                    </xsl:if> <!--第二層 -->
                </li>
            </xsl:for-each> <!--第一層 -->
        </ol>
    </xsl:template> <!-- 網站導覽 --> <!--  路徑連結  -->
    <xsl:template match="hpMain" mode="RssxPath">
        <div class="Path">
            <xsl:value-of select="//xxpath" />
            <a>
                <xsl:attribute name="href">dp.asp?mp=<xsl:value-of select="//Mp" /></xsl:attribute>
                <xsl:attribute name="title">
                    <xsl:value-of select="//xxhome" />
                </xsl:attribute>
                <xsl:value-of select="//xxhome" />
            </a>
            <img>
                <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/path_arrow.gif</xsl:attribute>
                <xsl:attribute name="alt">
                    <xsl:value-of select="//xxarrow" />
                </xsl:attribute>
            </img>
            <xsl:value-of select="RSSTop/Caption" />
        </div>
    </xsl:template> <!--  路徑連結  --> <!--  Rss 網站地圖資訊／開始  -->
    <xsl:template match="hpMain" mode="RSSTop">
        <h3>
            <xsl:value-of select="RSSTop/Article /Caption" />
        </h3>
        <xsl:value-of disable-output-escaping="yes" select="RSSTop/Article/ContentAll" />
    </xsl:template> <!--  Rss 網站地圖資訊／結束  --> <!-- Rss 網站導覽／開始 -->
    <xsl:template match="hpMain" mode="RssSitemap">
        <h3>
            <xsl:value-of select="RSSTop/Caption" /> <!--xsl:value-of select="//CtUnitName"/--></h3>
        <ul> <!--第一層 -->
            <xsl:for-each select="Sitemap/Menucat[@dataparent='0']">
                <xsl:variable name="thisPosition">
                    <xsl:value-of select="position()" />
                </xsl:variable>
                <li>
                    <xsl:choose>
                        <xsl:when test="xURL!=''"> <!--xsl:value-of select="$thisPosition" />.-->
                            <a>
                                <xsl:attribute name="href">
                                    <xsl:value-of select="xURL" />
                                </xsl:attribute>
                                <xsl:attribute name="title">
                                    <xsl:value-of select="Caption" />
                                </xsl:attribute>
                                <xsl:if test="xURL/@newWindow='Y'">
                                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                                </xsl:if>
                                <xsl:value-of select="Caption" />
                            </a>
                            <xsl:if test="YNrss='Y'">
                                <a>
                                    <xsl:attribute name="href">ws_RSS.asp?ctNodeId=<xsl:value-of select="@iCuItem" /></xsl:attribute>
                                    <img>
                                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/rss.gif</xsl:attribute>
                                        <xsl:attribute name="alt">
                                            <xsl:value-of select="//xxrss" />
                                        </xsl:attribute>
                                    </img>
                                </a>
                            </xsl:if>
                        </xsl:when>
                        <xsl:otherwise> <!--xsl:value-of select="$thisPosition" />.-->
                            <xsl:value-of select="Caption" />
                        </xsl:otherwise>
                    </xsl:choose> <!--第二層 -->
                    <xsl:variable name="thisCaptionID">
                        <xsl:value-of select="@iCuItem" />
                    </xsl:variable>
                    <xsl:if test="@isUnit='C'">
                        <ul>
                            <xsl:for-each select="ancestor::hpMain//Sitemap/Menucat[@dataparent=$thisCaptionID]">
                                <li>
                                    <xsl:variable name="thisPosition2">
                                        <xsl:value-of select="position()" />
                                    </xsl:variable>
                                    <xsl:choose>
                                        <xsl:when test="xURL!=''"> <!--xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> .-->
                                            <a>
                                                <xsl:attribute name="href">
                                                    <xsl:value-of select="xURL" />
                                                </xsl:attribute>
                                                <xsl:attribute name="title">
                                                    <xsl:value-of select="Caption" />
                                                </xsl:attribute>
                                                <xsl:if test="xURL/@newWindow='Y'">
                                                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                                                </xsl:if>
                                                <xsl:value-of select="Caption" />
                                            </a>
                                            <xsl:if test="YNrss='Y'">
                                                <a>
                                                    <xsl:attribute name="href">ws_RSS.asp?ctNodeId=<xsl:value-of select="@iCuItem" /></xsl:attribute>
                                                    <img>
                                                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/rss.gif</xsl:attribute>
                                                        <xsl:attribute name="alt">
                                                            <xsl:value-of select="//xxrss" />
                                                        </xsl:attribute>
                                                    </img>
                                                </a>
                                            </xsl:if>
                                        </xsl:when>
                                        <xsl:otherwise> <!--xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> .-->
                                            <xsl:value-of select="Caption" />
                                        </xsl:otherwise>
                                    </xsl:choose> <!--第三層 -->
                                    <xsl:variable name="thisCaptionID2">
                                        <xsl:value-of select="@iCuItem" />
                                    </xsl:variable>
                                    <xsl:if test="@isUnit='C'">
                                        <ul>
                                            <xsl:for-each select="ancestor::hpMain//Sitemap/Menucat[@dataparent=$thisCaptionID2]">
                                                <li>
                                                    <xsl:variable name="thisPosition3">
                                                        <xsl:value-of select="position()" />
                                                    </xsl:variable>
                                                    <xsl:choose>
                                                        <xsl:when test="xURL!=''"> <!--xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> - <xsl:value-of select="$thisPosition3" /> .-->
                                                            <a>
                                                                <xsl:attribute name="href">
                                                                    <xsl:value-of select="xURL" />
                                                                </xsl:attribute>
                                                                <xsl:attribute name="title">
                                                                    <xsl:value-of select="Caption" />
                                                                </xsl:attribute>
                                                                <xsl:if test="xURL/@newWindow='Y'">
                                                                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                                                                </xsl:if>
                                                                <xsl:value-of select="Caption" />
                                                            </a>
                                                            <xsl:if test="YNrss='Y'">
                                                                <a>
                                                                    <xsl:attribute name="href">ws_RSS.asp?ctNodeId=<xsl:value-of select="@iCuItem" /></xsl:attribute>
                                                                    <img>
                                                                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/rss.gif</xsl:attribute>
                                                                        <xsl:attribute name="alt">
                                                                            <xsl:value-of select="//xxrss" />
                                                                        </xsl:attribute>
                                                                    </img>
                                                                </a>
                                                            </xsl:if>
                                                        </xsl:when>
                                                        <xsl:otherwise> <!--xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> - <xsl:value-of select="$thisPosition3" /> .-->
                                                            <xsl:value-of select="Caption" />
                                                        </xsl:otherwise>
                                                    </xsl:choose> <!--第四層 -->
                                                    <xsl:variable name="thisCaptionID3">
                                                        <xsl:value-of select="@iCuItem" />
                                                    </xsl:variable>
                                                    <xsl:if test="@isUnit='C'">
                                                        <ul>
                                                            <xsl:for-each select="ancestor::hpMain//Sitemap/Menucat[@dataparent=$thisCaptionID3]">
                                                                <li>
                                                                    <xsl:variable name="thisPosition4">
                                                                        <xsl:value-of select="position()" />
                                                                    </xsl:variable>
                                                                    <xsl:choose>
                                                                        <xsl:when test="xURL!=''"> <!--xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> -
															<xsl:value-of select="$thisPosition3" /> - <xsl:value-of select="$thisPosition4" /> .-->
                                                                            <a>
                                                                                <xsl:attribute name="href">
                                                                                    <xsl:value-of select="xURL" />
                                                                                </xsl:attribute>
                                                                                <xsl:attribute name="title">
                                                                                    <xsl:value-of select="Caption" />
                                                                                </xsl:attribute>
                                                                                <xsl:if test="xURL/@newWindow='Y'">
                                                                                    <xsl:attribute name="target">_nwGip</xsl:attribute>
                                                                                </xsl:if>
                                                                                <xsl:value-of select="Caption" />
                                                                            </a>
                                                                            <xsl:if test="YNrss='Y'">
                                                                                <a>
                                                                                    <xsl:attribute name="href">ws_RSS.asp?ctNodeId=<xsl:value-of select="@iCuItem" /></xsl:attribute>
                                                                                    <img>
                                                                                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/rss.gif</xsl:attribute>
                                                                                        <xsl:attribute name="alt">
                                                                                            <xsl:value-of select="//xxrss" />
                                                                                        </xsl:attribute>
                                                                                    </img>
                                                                                </a>
                                                                            </xsl:if>
                                                                        </xsl:when>
                                                                        <xsl:otherwise> <!--xsl:value-of select="$thisPosition" /> - <xsl:value-of select="$thisPosition2" /> -
															<xsl:value-of select="$thisPosition3" /> - <xsl:value-of select="$thisPosition4" /> .-->
                                                                            <xsl:value-of select="Caption" />
                                                                        </xsl:otherwise>
                                                                    </xsl:choose>
                                                                </li>
                                                            </xsl:for-each>
                                                        </ul>
                                                    </xsl:if> <!--第四層 -->
                                                </li>
                                            </xsl:for-each>
                                        </ul>
                                    </xsl:if> <!--第三層 -->
                                </li>
                            </xsl:for-each>
                        </ul>
                    </xsl:if> <!--第二層 -->
                </li>
            </xsl:for-each> <!--第一層 -->
        </ul>
    </xsl:template> <!-- Rss 網站導覽／結束 --> <!-- Accesskey C／開始 -->
    <xsl:template match="hpMain" mode="c">
        <a title="主要內容區" class="Accesskey" accesskey="C">
            <xsl:apply-templates select="." mode="link" />
        </a>
    </xsl:template> <!-- Accesskey C／結束 --> <!-- Accesskey M／開始 -->
    <xsl:template match="hpMain" mode="m">
        <a title="主選單" class="Accesskey" accesskey="M">
            <xsl:apply-templates select="." mode="link" />
        </a>
    </xsl:template> <!-- Accesskey M／結束 --> <!-- Accesskey R／開始 -->
    <xsl:template match="hpMain" mode="r">
        <a title="右欄" class="Accesskey" accesskey="R">
            <xsl:apply-templates select="." mode="link" />
        </a>
    </xsl:template> <!-- Accesskey R／結束 --> <!-- Accesskey S／開始 -->
    <xsl:template match="hpMain" mode="s">
        <a title="Search" class="Accesskey" accesskey="S">
            <xsl:apply-templates select="." mode="link" />
        </a>
    </xsl:template> <!-- Accesskey S／結束 --> <!-- Accesskey U／開始 -->
    <xsl:template match="hpMain" mode="u">
        <a title="網站標題及導覽列" class="Accesskey" accesskey="U">
            <xsl:apply-templates select="." mode="link" />
        </a>
    </xsl:template> <!-- Accesskey U／結束 --> <!-- Link／開始 -->
    <xsl:template match="hpMain" mode="link">
	<xsl:attribute name="href">
		content.asp?mp=<xsl:value-of select="//Mp" />&amp;CuItem=<xsl:value-of select="//SitemapTop /Article/@iCuItem" />
	</xsl:attribute>:::
</xsl:template> <!-- Link／結束 --> <!-- 文章／開始 -->
    <xsl:template match="MainArticle">
        <div class="Article">
            <h1>
                <xsl:value-of select="Caption" />
            </h1>
            <div class="Date">
                張貼日期：<xsl:value-of select="//PostDate" />
            </div>
            <div class="Body">
                <xsl:if test="xImgFile">
                    <div class="Image">
                        <a target="_blank">
                            <xsl:attribute name="href">
                                /<xsl:value-of select="//MainArticle/xImgFile" />
                            </xsl:attribute>
                            <xsl:attribute name="title">
                                <xsl:value-of select="//MainArticle/Caption" />
                            </xsl:attribute>
                            <img border="0">
                                <xsl:attribute name="src">/<xsl:value-of select="//MainArticle/xImgFile" />
                                </xsl:attribute>
                                <xsl:attribute name="alt">
                                    <xsl:value-of select="//MainArticle/Caption" />
                                </xsl:attribute>
                            </img>
                        </a>
                        <p>
                            <xsl:value-of select="//MainArticle/Caption" />
                        </p>
                    </div>
                </xsl:if>
                <p>
                    <xsl:value-of disable-output-escaping="yes" select="//MainArticle/Content" />
                </p>
                <xsl:if test="//AttachmentList">
                    <xsl:apply-templates select="//AttachmentList" />
                </xsl:if>
            </div>
			<xsl:apply-templates select="//pHTML" />
            <div class="Foot"></div>
        </div>
    </xsl:template> <!-- 文章／結束 --> <!-- 附件列表／開始-->
	<xsl:template match="//pHTML">
		<xsl:copy-of select="." />
	</xsl:template>
    <xsl:template match="AttachmentList">
        <h3>
            <xsl:value-of select="//xxattachment" />
        </h3>
        <ul class="Download">
            <xsl:for-each select="//AttachmentList/Attachment">
                <li>
                    <a target="_nwGIP">
                        <xsl:attribute name="href">
                            <xsl:value-of select="URL" />
                        </xsl:attribute>
                        <xsl:attribute name="title">
                            <xsl:value-of select="Caption" />
                        </xsl:attribute>
                        <xsl:value-of select="Caption" />
                    </a>
                </li>
            </xsl:for-each>
        </ul>
    </xsl:template> <!-- 附件列表／結束 --> <!-- 相關聯結／開始 -->
    <xsl:template match="ReferenceList">
        <div class="Relation">
            <h3>
                <xsl:value-of select="//xxreference" />
            </h3>
            <ol>
                <xsl:for-each select="//ReferenceList/Reference">
                    <li>
                        <span class="Title">
                            <a>
                                <xsl:attribute name="href">
                                    <xsl:value-of select="URL" />
                                </xsl:attribute>
                                <xsl:attribute name="title">
                                    <xsl:value-of select="Caption" />
                                </xsl:attribute>
                                <xsl:value-of select="Caption" />
                            </a>
                        </span>
                        <span class="Date">
                            <xsl:value-of select="PostDate" />
                        </span>
                        <span class="Brief">
                            <xsl:value-of select="Content" />
                        </span>
                    </li>
                </xsl:for-each>
            </ol>
        </div>
    </xsl:template> <!-- 相關聯結／結束 -->
</xsl:stylesheet>
