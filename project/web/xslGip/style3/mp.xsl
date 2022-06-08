<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
xmlns:hyweb="urn:gip-hyweb-com" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
<xsl:include href="info.xsl"/>

<xsl:template match="hpMain">
	<html xmlns="http://www.w3.org/1999/xhtml">
		<head>
			<xsl:apply-templates select="." mode="Header"/>
		</head>
		<script language="javascript" type="text/javascript">
			function bodyOnload() {
			    loadtagcloud();
				clickButton('A');
				clickKnowledgeButton('KA');
			}

			function loadtagcloud() {
			var url="/services/tagCloud.aspx?cmd=gettagcloud";
			var myAjax = new Ajax(url, {
					method: 'get',
					update: $('tagcloud')
			}).request();
         }
		</script>
		<body onload="bodyOnload();">
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
					<xsl:apply-templates select="BlockA" />
					<!--List Star (style D)-->
					<xsl:apply-templates select="BlockB" />
					<!--List End-->
					<xsl:apply-templates select="BlockD" />

					<xsl:apply-templates select="KMDOC" />

					<!--hotessay End-->
				 </td>
				<td class="right">
					<xsl:apply-templates select="." mode="r"/>
					<!--search star (style C)-->
					<xsl:apply-templates select="." mode="xsSearch"/>
					<!--search end-->
					<!--主題報導專欄login star-->
					<xsl:apply-templates select="BlockC" />
					<!--主題報導專欄 end-->
					<!--login star-->
					<xsl:apply-templates select="." mode="xsMember"/>
					<!--login end-->
					<!--epaper star (style B)-->
					<xsl:apply-templates select="." mode="xsePaper"/>
					<!--epaper end-->
					<!--tag Cloud -->
					<div class="TagCloud">
					    <h2>熱門關鍵字</h2>
					    <div id ="tagcloud">
						</div>
					</div>
					<!--ag Cloud end-->
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


<!-- 最燒推薦主題館 Start -->
<xsl:template match="BlockA">
	<h3><xsl:value-of select="Caption" /></h3>
	<!--sublayout star (style C)-->
	<table summary="內層排版表格" class="sublayout">
		<tr>
			<xsl:for-each select="Article">
				<xsl:if test="position() mod 2= 1">
					<xsl:text disable-output-escaping="yes"><![CDATA[<td class="col1" width="50%">]]></xsl:text>
				</xsl:if>
				<xsl:if test="position() mod 2= 0">
					<xsl:text disable-output-escaping="yes"><![CDATA[<td class="col2">]]></xsl:text>
				</xsl:if>
				<h4>
        <a>
					<xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute>
					<xsl:attribute name="target">_blank</xsl:attribute>
					<xsl:if test="xImgFile">
					<img>
						<xsl:attribute name="src"><xsl:value-of select="xImgFile" /></xsl:attribute>
						<xsl:attribute name="alt"><xsl:value-of select="Caption" /></xsl:attribute>
					</img>
					</xsl:if>
					<xsl:value-of select="Caption" />
				</a>
				</h4>
				<p>
					<xsl:value-of disable-output-escaping="yes" select="substring(ContentText,1,85)" />...
				<a>
					<xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute>
					<xsl:attribute name="target">_blank</xsl:attribute>
					前往
				</a>
				</p>
				<xsl:text disable-output-escaping="yes"><![CDATA[</td>]]></xsl:text>
			</xsl:for-each>
		</tr>
	</table>
</xsl:template>
<!-- 最燒推薦主題館 End -->


<!-- 最新主題館文章 Start -->
<xsl:template match="BlockB">
	<div class="topic">
		<h3><xsl:value-of select="Caption" /></h3>
		<xsl:for-each select="Article">
		<div>
			<h4><xsl:value-of select="Caption" /><span><xsl:value-of select="PostDate" /></span></h4>
			<p>
				資料來源：<xsl:value-of disable-output-escaping="yes" select="substring(Content,1,100)" />
				<!--a title="詳全文"><xsl:attribute name="href">ct.asp?xItem=<xsl:value-of select="@iCuItem" />&amp;ctNode=<xsl:value-of select="../@xNode" /></xsl:attribute>
				詳全文</a-->
        <xsl:choose>
					<xsl:when test="contains(xURL,'http')">
						<a target="_blank" title="閱讀全文"><xsl:attribute name="href"><xsl:value-of select="xURL" /></xsl:attribute>閱讀全文</a>
					</xsl:when>
					<xsl:when test="contains(xURL,'Knowledge')">
						<a target="_blank" title="閱讀全文"><xsl:attribute name="href"><xsl:value-of select="xURL" /></xsl:attribute>閱讀全文</a>
					</xsl:when>
					<xsl:otherwise>
					<a target="_blank" title="閱讀全文"><xsl:attribute name="href">/subject/<xsl:value-of select="xURL" /></xsl:attribute>閱讀全文</a>
					</xsl:otherwise>
        </xsl:choose>
			</p>
		</div>
		</xsl:for-each>
	</div>
</xsl:template>
<!-- 最新主題館文章 End -->

<!-- 主題報導專欄 Start -->
<xsl:template match="BlockC">
		<div class="theme">
			<h2><xsl:value-of select="Caption" /></h2>
			<div class="body">
				<xsl:for-each select="Article">
					<xsl:if test="xImgFile">
						<a>
							<xsl:attribute name="href"><xsl:value-of select="xImgFile" /></xsl:attribute>
							<xsl:attribute name="target">_blank</xsl:attribute>
							<img border="0">
								<xsl:attribute name="src"><xsl:value-of select="xImgFile" /></xsl:attribute>
								<xsl:attribute name="alt"><xsl:value-of select="Caption" /></xsl:attribute>
							</img>
						</a>
					</xsl:if>
					<p>
						<a>
							<xsl:attribute name="href">ct.asp?xItem=<xsl:value-of select="@iCuItem" />&amp;ctNode=<xsl:value-of select="../@xNode" /></xsl:attribute>
							<xsl:attribute name="title"><xsl:value-of disable-output-escaping="yes" select="substring(Content,1,30)" /></xsl:attribute>
							<xsl:value-of disable-output-escaping="yes" select="substring(Content,1,30)" />
						</a>...
					</p>
				</xsl:for-each>
			</div>
			<div class="foot">&amp;nbsp;</div>
		</div>
</xsl:template>
<!-- 主題報導專欄 End -->


<!-- 最新消息 Start -->
<xsl:template match="BlockD">
	<div class="hotQa">
		<h3><xsl:value-of select="Caption" /></h3>
		<ul class="list">
			<xsl:for-each select="Article">
				<li>
					<a>
						<xsl:choose>
							<xsl:when test="@ShowType='3'">
								<xsl:attribute name="href"><xsl:value-of select="FileDownLoad"/></xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="href"><xsl:value-of select="xURL"/></xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
						<xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute>
						<xsl:if test="@newWindow='Y'"><xsl:attribute name="target">_blank</xsl:attribute></xsl:if>
						<xsl:value-of select="Caption" />
					</a>
					<span class="from"><xsl:value-of select="DeptName" /></span>
					<span class="date"><xsl:value-of disable-output-escaping="yes" select="PostDate" /></span>
				</li>
			</xsl:for-each>
		</ul>
		<!--more star (style A)-->
		<div class="more">
			<a>
				<xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="@xNode" />&amp;mp=<xsl:value-of select="//mp" /></xsl:attribute>
				<xsl:attribute name="title">more</xsl:attribute>more
			</a>
		</div>
		<!--more end-->
	</div>
</xsl:template>
<!-- 最新消息 End -->

<!-- 熱門知識文件 Start -->
<xsl:template match="KMDOC">

<div class="hotQa">
		<h3>熱門知識家問答</h3>
		<script type="text/javascript" src="mootools.v1.11.js"></script>
		<script language="javascript">

		function clickKnowledgeButton(value){
			clearKnowledgeButtonClass();
			document.getElementById(value).className = "now";
			KnowledgeAjax(value);
		}

		function clearKnowledgeButtonClass(){
			document.getElementById("KA").className="";
			document.getElementById("KB").className="";
			document.getElementById("KC").className="";
		}

		function KnowledgeAjax(value){
			var url = "knowledgesql.asp?type="+value;
			var myAjax = new Ajax(url, {
					method: 'get',
					update: $('KnowledgeId')
			}).request();
		}
		</script>
			<!--div class="rss">RSS訂閱</div-->
				<ul class="function">
					<li id="KA" class="now"><a href="javascript:clickKnowledgeButton('KA')">最新發問</a></li>
					<li id="KB"><a href="javascript:clickKnowledgeButton('KB')">熱門討論</a></li>
					<li id="KC"><a href="javascript:clickKnowledgeButton('KC')">專家補充</a></li>
				</ul>
			<ul class="list"><div id="KnowledgeId" ></div></ul>
		<div class="more"><a href="/knowledge/knowledge_lp.aspx?ArticleType=A&amp;CategoryId=">more...</a></div>
		<!--more end-->
	</div>

	<div class="hotessay">
		<h3>熱門知識庫文章</h3>
		<script type="text/javascript" src="mootools.v1.11.js"></script>
		<script language="javascript">

      function clickButton(value){
      clearButtonClass();
      document.getElementById(value).className = "now";
      testAjax(value);
      }

      function clearButtonClass(){
      document.getElementById("A").className="";
      document.getElementById("B").className="";
      }

      function testAjax(value){
      var send = "cmd={'id':'gethotdocuments','key':'" + value + "'}";
      var url="ajaxservice.aspx?"+send;
      var myAjax = new Ajax(url, {
      method: 'get',
      update: $('tokenId')
      }).request();
      }
    </script>
				<ul class="function">
					<li id="A" class="now"><a href="javascript:clickButton('A')">最新文章</a></li>
					<li id="B"><a href="javascript:clickButton('B')">最多人瀏覽</a></li>
				</ul>
			<ul class="list"><div id="tokenId" ></div></ul>
		<div class="more"><a href="/category/categorylist.aspx?CategoryId=&amp;ActorType=000&amp;t=a">more...</a></div>
		<!--more end-->
	</div>

</xsl:template>
<!-- 熱門知識文件 End -->

</xsl:stylesheet>
