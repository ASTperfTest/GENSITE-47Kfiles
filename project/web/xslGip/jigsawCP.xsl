<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
xmlns:user="urn:user-namespace-here" version="1.0">
  <xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />

  <xsl:include href="info.xsl"/>

  <xsl:template match="hpMain">
    <html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
      <head>
        <xsl:apply-templates select="." mode="Header"/>
		<script language="JavaScript">
			function toggleDiscussionSection(){
				var elem = document.getElementById('DiscussionBox');
				if(elem.style.display == "none")
				{elem.style.display="";}
				else{elem.style.display="none";}
			}
			function checkNull()
			{
				var txtDiscussion=document.getElementById('txtDiscussion');
				if(txtDiscussion.value=="")
				{
					alert('請輸入留言內容');
					return false;
				}
				else
				{return true;}
			}
		</script>
      </head>

      <body>
        <!--header star (style A)-->
        <xsl:apply-templates select="." mode="RunActive"/>
        <!--header End-->
        <!--nav star (style A)-->
        <xsl:apply-templates select="." mode="toplink"/>
        <!--nav end-->
        <!--subnav star (style G)-->
        <xsl:apply-templates select="." mode="weblink"/>
        <!--subnav end-->

        <!--layout Table star (style A)-->
        <table class="layout">
          <tr>
            <td class="left">
              <xsl:apply-templates select="." mode="l"/>
              <!--season star-->
              <xsl:apply-templates select="." mode="SolarTerms"/>
              <!--season end-->
              <!--menu star (style C)-->
              <xsl:apply-templates select="." mode="Menu"/>
              <!--menu end-->

              <!--RSS star-->
              <xsl:apply-templates select="." mode="xsRSS"/>
              <!--RSS end-->
              <!--ad star style A-->
              <xsl:apply-templates select="." mode="xsAD"/>
              <!--ad end-->
            </td>
            <td class="center">
              <xsl:apply-templates select="." mode="c"/>
              <!--path star (style B)-->
              <xsl:apply-templates select="." mode="xsPath"/>
              <!--path end-->
              <xsl:apply-templates select="." mode="jigsaw"/>
              <!--hotessay End-->
            </td>
            <td class="right">
              <xsl:apply-templates select="." mode="r"/>
              <!--search star (style C)-->
              <xsl:apply-templates select="." mode="xsSearch"/>
              <!--search end-->
              <!--login star-->
              <xsl:apply-templates select="." mode="xsMember"/>
              <!--login end-->
              <!--epaper star (style B)-->
              <xsl:apply-templates select="." mode="xsePaper"/>
              <!--epaper end-->
            </td>
          </tr>
        </table>
        <!--layout Table End-->
        <!--footer star (style A)-->
        <div class="footer">
          <xsl:apply-templates select="." mode="xsFooter"/>
          <xsl:apply-templates select="." mode="xsCopyright"/>
        </div>
        <!--footer End-->
      </body>
    </html>
  </xsl:template>

  <xsl:template match="hpMain" mode="jigsaw">
    <!-- 內容區 Start -->
    <div class="jigsaw">
      <div class="head"></div>
      <div class="body">
        <h2>農業知識拼圖</h2>
        <!-- 最新議題：開始 -->
        <xsl:if test="jigsaw/latestissue/isopen='Y'">
          <div class="jigsawnew">
            <h5>最新議題</h5>
            <xsl:for-each select="jigsaw/latestissue/article">
              <xsl:if test="top='Y'">
                <div class="content">
                  <xsl:if test="img!=''">
                    <img>
                      <xsl:attribute name="src">
                        /public/Data/<xsl:value-of select="img" />
                      </xsl:attribute>
                      <xsl:attribute name="alt">
                        <xsl:value-of select="stitle" />
                      </xsl:attribute>
                    </img>
                  </xsl:if>
                  <h6>
                    <a>
                      <xsl:attribute name="href">
                        <xsl:value-of select="path" />
                      </xsl:attribute>
                      <xsl:value-of select="stitle" />
                    </a>
                  </h6>
                  <p>
                    <xsl:value-of select="content" />
                    <a target="_nwGip">
                      <xsl:attribute name="href">
                        <xsl:value-of select="path" />
                      </xsl:attribute>
                      詳全文
                    </a>
                  </p>
                </div>
              </xsl:if>
            </xsl:for-each>
            <ol>
              <xsl:for-each select="jigsaw/latestissue/article">
                <xsl:if test="top='N'">
                  <li>
                    <a target="_nwGip">
                      <xsl:attribute name="href">
                        <xsl:value-of select="path" />
                      </xsl:attribute>
                      <xsl:value-of select="stitle" />
                    </a>
                  </li>
                </xsl:if>
              </xsl:for-each>
            </ol>
          </div>
        </xsl:if>
        <!-- 最新議題：結束 -->
        <xsl:if test="jigsaw/issuearticle/isopen='Y'">
          <h5>議題關聯知識文章</h5>
          <xsl:choose>
            <xsl:when test="jigsaw/issuearticle/order/unit='1'">
              <xsl:apply-templates select="." mode="unit" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/subject='1'">
              <xsl:apply-templates select="." mode="subject" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/tank='1'">
              <xsl:apply-templates select="." mode="tank" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/home='1'">
              <xsl:apply-templates select="." mode="home" />
            </xsl:when>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="jigsaw/issuearticle/order/unit='2'">
              <xsl:apply-templates select="." mode="unit" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/subject='2'">
              <xsl:apply-templates select="." mode="subject" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/tank='2'">
              <xsl:apply-templates select="." mode="tank" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/home='2'">
              <xsl:apply-templates select="." mode="home" />
            </xsl:when>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="jigsaw/issuearticle/order/unit='3'">
              <xsl:apply-templates select="." mode="unit" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/subject='3'">
              <xsl:apply-templates select="." mode="subject" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/tank='3'">
              <xsl:apply-templates select="." mode="tank" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/home='3'">
              <xsl:apply-templates select="." mode="home" />
            </xsl:when>
          </xsl:choose>
          <xsl:choose>
            <xsl:when test="jigsaw/issuearticle/order/unit='4'">
              <xsl:apply-templates select="." mode="unit" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/subject='4'">
              <xsl:apply-templates select="." mode="subject" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/tank='4'">
              <xsl:apply-templates select="." mode="tank" />
            </xsl:when>
            <xsl:when test="jigsaw/issuearticle/order/home='4'">
              <xsl:apply-templates select="." mode="home" />
            </xsl:when>
          </xsl:choose>
          <xsl:apply-templates select="." mode="media" />
        </xsl:if>

        <!-- 資源推薦：開始 -->
        <xsl:if test="jigsaw/reflink/isopen='Y'">
        <xsl:if test="count(jigsaw/reflink/article) >0">
          <h5>最佳資源推薦</h5>
            <table class="jigsawtype03" summary="結果條列式">
              <tr>
                <th colspan="3" scope="col">推薦超聯結列表</th>
              </tr>
              <xsl:for-each select="jigsaw/reflink/article">
                <tr>
                  <td width="5%" align="center">
                    <xsl:value-of select="position()"/>.
                  </td>
                  <td width="15%" align="center">
                    <xsl:value-of select="xpostdate" />
                  </td>
                  <td>
                    <a target="_nwGip">
                      <xsl:attribute name="href">
                        <xsl:value-of select="url" />
                      </xsl:attribute>
                      <xsl:value-of select="title" />
                    </a>
                  </td>
                </tr>
              </xsl:for-each>
            </table>
        </xsl:if>
        </xsl:if>
        <!-- 資源推薦：結束 -->
		<h5>議題分享</h5>
            <table class="jigsawtype03" summary="結果條列式">
              <tr>
                <th colspan="2" scope="col">留言</th>
				<td width="20%">					
					<a href="#" onclick="toggleDiscussionSection();return false;">
                      我要分享
                    </a>
                </td>
              </tr>			  
              <xsl:for-each select="jigsaw/discussion/article">
                <tr>
                  <td width="5%" align="center">
                    <xsl:value-of select="position()"/>.
                  </td>
                  <td width="15%" align="center">
                    <xsl:value-of select="iEditor" /> 發表於 <xsl:value-of select="xpostdate" />
                  </td>
                  <td>                    
                      <xsl:value-of select="xBody" />                    
                  </td>
                </tr>
              </xsl:for-each>
			  <tr >
				<td colspan="3">				
					<form name="formDiscussion" method="post" action="/addDiscussion.asp" id="formDiscussion">
						<xsl:attribute name="action">
							/addDiscussion.asp?xItem=<xsl:value-of select="//hpMain/queryItems/xItem" />
						</xsl:attribute>
					  
						<div id="DiscussionBox" style="display:none">			
							<textarea name="txtDisCussion" rows="10" cols="60" id="txtDisCussion"></textarea><br/>
							請輸入圖片驗證碼：<image src="/VerifyImage.asp" width="80pt" style="border-bottom:1px solid #999999;border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999;"/><input type="text" name="CheckCode"/>
							<input type="submit" name="ButtonSave" value="確認" id="ButtonSave" onclick="return checkNull();"/>
							<input type="button" name="ButtonCancle" value="取消" id="ButtonCancle" />							
						</div>
					</form>
				</td>				
			  </tr>
            </table>
      </div>
	  
      <div class="foot"></div>
      <div class="jigsawtop">
        <a href="#">top</a>
      </div>
    </div>
    <!-- 內容區 End -->
  </xsl:template>

  <xsl:template match="hpMain" mode="unit">
    <xsl:if test="jigsaw/issuearticle/unit/count > 0">
      <table class="jigsawtype02" summary="結果條列式">
        <tr>
          <th colspan="3" scope="col">入口網關聯文章</th>
        </tr>
        <xsl:for-each select="jigsaw/issuearticle/unit/article">
          <tr>
            <td width="5%" align="center">
              <xsl:value-of select="position()"/>.
            </td>
            <td width="15%" align="center">
              <xsl:value-of select="xpostdate" />
            </td>
            <td>
              <a target="_nwGip">
                <xsl:attribute name="href">
                  <xsl:value-of select="path" />
                </xsl:attribute>
                <xsl:value-of select="stitle" />
              </a>
            </td>
          </tr>
        </xsl:for-each>
      </table>
    </xsl:if>
  </xsl:template>

  <xsl:template match="hpMain" mode="subject">
    <xsl:if test="jigsaw/issuearticle/subject/count > 0">
      <table class="jigsawtype03" summary="結果條列式">
        <tr>
          <th colspan="3" scope="col">主題館關聯文章</th>
        </tr>
        <xsl:for-each select="jigsaw/issuearticle/subject/article">
          <tr>
            <td width="5%" align="center">
              <xsl:value-of select="position()"/>.
            </td>
            <td width="15%" align="center">
              <xsl:value-of select="xpostdate" />
            </td>
            <td>
              <a target="_nwGip">
                <xsl:attribute name="href">
                  <xsl:value-of select="path" />
                </xsl:attribute>
                <xsl:value-of select="stitle" />
              </a>
            </td>
          </tr>
        </xsl:for-each>
      </table>
    </xsl:if>
  </xsl:template>

  <xsl:template match="hpMain" mode="tank">
    <xsl:if test="jigsaw/issuearticle/tank/count > 0">
      <table class="jigsawtype04" summary="結果條列式">
        <tr>
          <th colspan="3" scope="col">知識庫關聯文章</th>
        </tr>
        <xsl:for-each select="jigsaw/issuearticle/tank/article">
          <tr>
            <td width="5%" align="center">
              <xsl:value-of select="position()"/>.
            </td>
            <td width="15%" align="center">
              <xsl:value-of select="xpostdate" />
            </td>
            <td>
              <a target="_nwGip">
                <xsl:attribute name="href">
                  <xsl:value-of select="path" />
                </xsl:attribute>
                <xsl:value-of select="stitle" />
              </a>
            </td>
          </tr>
        </xsl:for-each>
      </table>
    </xsl:if>
  </xsl:template>

  <xsl:template match="hpMain" mode="home">
    <xsl:if test="jigsaw/issuearticle/home/count > 0">
      <table class="jigsawtype05" summary="結果條列式">
        <tr>
          <th colspan="3" scope="col">知識家關聯討論</th>
        </tr>
        <xsl:for-each select="jigsaw/issuearticle/home/article">
          <tr>
            <td width="5%" align="center">
              <xsl:value-of select="position()"/>.
            </td>
            <td width="15%" align="center">
              <xsl:value-of select="xpostdate" />
            </td>
            <td>
              <a target="_nwGip">
                <xsl:attribute name="href">
                  <xsl:value-of select="path" />
                </xsl:attribute>
                <xsl:value-of select="stitle" />
              </a>
            </td>
          </tr>
        </xsl:for-each>
      </table>
    </xsl:if>
  </xsl:template>

  <xsl:template match="hpMain" mode="media">
    <xsl:if test="jigsaw/issuearticle/media/count > 0">
      <table class="jigsawtype05" summary="結果條列式">
        <tr>
          <th colspan="3" scope="col">議題關聯影音</th>
        </tr>
        <xsl:for-each select="jigsaw/issuearticle/media/article">
          <tr>
            <td width="5%" align="center">
              <xsl:value-of select="position()"/>.
            </td>
            <td width="15%" align="center">
              <xsl:value-of select="xpostdate" />
            </td>
            <td>
              <a target="_nwGip">
                <xsl:attribute name="href">
                  <xsl:value-of select="path" />
                </xsl:attribute>
                <xsl:value-of select="stitle" />
              </a>
            </td>
          </tr>
        </xsl:for-each>
      </table>
    </xsl:if>
  </xsl:template>

</xsl:stylesheet>

