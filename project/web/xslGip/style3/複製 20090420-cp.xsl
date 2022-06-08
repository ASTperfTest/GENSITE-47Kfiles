<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" />
<xsl:include href="info.xsl"/>

<xsl:template match="hpMain">
	<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
		<head>
			<xsl:apply-templates select="." mode="Header"/>
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
					
					<xsl:apply-templates select="." mode="textSearch"/>
					
					<!--path end-->
					<h3><xsl:value-of select="//xPath/UnitName"/></h3>
					<form method="post">
					<input type="hidden" name="searchText" value="" />
					<xsl:apply-templates select="." mode="xsFunction"/>
					<!-- 最新發問TAB LIST／開始 -->
					<!-- 內容區 Start -->
					<xsl:choose>
						<xsl:when test="//cpStyle='A'">
							<xsl:apply-templates select="//MainArticle" mode="styleA" />
						</xsl:when>
						<xsl:when test="//cpStyle='B'">
							<xsl:apply-templates select="//MainArticle" mode="styleB" />
						</xsl:when>
						<xsl:when test="//cpStyle='C'">
							<xsl:apply-templates select="//MainArticle" mode="styleC" />
						</xsl:when>
						<xsl:when test="//cpStyle='D'">
							<xsl:apply-templates select="//MainArticle" mode="styleD" />
						</xsl:when>
						<xsl:when test="//cpStyle='E'">
							<xsl:apply-templates select="//MainArticle" mode="styleE" />
						</xsl:when>
						<xsl:otherwise>
							<xsl:apply-templates select="//MainArticle" />
						</xsl:otherwise>
					</xsl:choose>							
					<!-- 內容區 結束 -->						
					</form>
					<!-- 最新發問TAB LIST／結束 -->
					<!-- 態度投票：開始 -->
					<xsl:apply-templates select="." mode="attitudeVote" />
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
					<xsl:apply-templates select="." mode="relatedDocument"/>
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

<!-- 中央主資料區 Start-->
<xsl:template match="MainArticle">
	<div id="Magazine">
		<div class="Event">
			<div id="cp">
				<div class="experts">
					<h4><xsl:value-of select="Caption" /></h4>
					<xsl:if test="xImgFile">
						<!--div class="exphoto"-->
						<div class="phoimg">
							<img>
								<xsl:attribute name="src"><xsl:value-of select="xImgFile" /></xsl:attribute>
								<xsl:attribute name="alt"><xsl:value-of select="Caption" /></xsl:attribute>
							</img>
						</div>
						<!--/div-->
					</xsl:if>
					<p><xsl:value-of disable-output-escaping="yes" select="Content" /></p>	
				</div>
        <!--download start (B)-->
				<xsl:apply-templates select="//AttachmentList" />
        <!--download end-->		
			</div>
		</div>
	</div>	
</xsl:template>	

<xsl:template match="AttachmentList">
	<div class="download">
		<h5>相關檔案下載</h5>
		<ul>
			<xsl:for-each select="Attachment">
				<li>
					<a target="_blank">
						<xsl:attribute name="href"><xsl:value-of select="URL" /></xsl:attribute>
						<xsl:attribute name="title"><xsl:value-of select="Caption" /></xsl:attribute>
						<xsl:value-of select="Caption" />
					</a>
				</li>
			</xsl:for-each>
		</ul>
	</div>
</xsl:template>
<!--download end-->

<!-- 以文找文 -->
<xsl:template match="hpMain" mode="textSearch">
	<div class="interact">
		<ul>
			<!-- <li class="ask"><a href="#">我要發問</a></li> -->
			<li class="find"><a href="javascript:doSearchReport();">以文找文</a></li>
		</ul>
	</div>
	<script language="JavaScript" type="text/JavaScript">
	<![CDATA[
	<!--
	function doSearchReport(){
	  var searchForm = document.forms[0];
	  searchForm.action="kp.asp?xdURL=Search/TextSearchAct.asp&mp=1";	  
	  searchForm.searchText.value=getSelectText();	 
	  if(""==searchForm.searchText.value){
	    alert("尚未選取查詢字!!");
	    return ;
	  }
	  if(getStrMomery(getSelectText())>750){
	    alert("選取中文字數不得超過250字!\n英文字不得超過700字!");
	    return ;
	  }
	  searchForm.submit();
	}

	function getStrMomery(str){
	  var wordAmount=0;
	  re   =   /[\u4E00-\u9FA0]/;   
	  for(var i=0;i<str.length;i++){
	    if(re.test(str.substr(i,1))){
	      wordAmount+=3;
	    }else{
	      wordAmount+=1;
	    }
	  } 
	  return wordAmount; 
	}

	function getSelectText(){
	  var desc="";
	  if(window.getSelection)
	    desc=window.getSelection();	  
	  if(document.getSelection)
	    desc=document.getSelection();	  
	  if(document.selection)
	    desc=document.selection.createRange().text;  
	  if(desc=="")
	    desc = getTextAreaSelect();	  
	  return desc;
	}

	function getTextAreaSelect(){
	  var textAreas = document.getElementsByTagName("textarea");
	  var str=""
	  for(var i=0;i<textAreas.length;i++){
	  var obj = textAreas[i];
	  obj.onselect=function(){
			for(var i=0;i<textAreas.length;i++){
				if(textAreas[i]!=this){
					textAreas[i].selectionStart = null;
					textAreas[i].selectionEnd = null;
				}
			}};	
	  if(obj.selectionEnd-obj.selectionStart>0){
		str = obj.value.substring(obj.selectionStart,obj.selectionEnd);
	        obj.focus();
	      }
	  }	  
	  return str;
	 }
	//-->
]]>
</script>
</xsl:template>	

<!-- 態度投票 -->
<xsl:template match="hpMain" mode="attitudeVote">
	
	<!-- 有登入 -->
	<xsl:if test="//login/status='true'">	
		<div class="toScore">
			<h5>這篇文章讓你覺得</h5>
			<form name="AttitudeVoteForm" action="/AttitudeVoteResult.asp" method="POST">
				<input type="hidden" name="xItem"><xsl:attribute name="value"><xsl:value-of select="//queryItems/xItem" /></xsl:attribute></input>
				<input type="hidden" name="ctNode"><xsl:attribute name="value"><xsl:value-of select="//queryItems/ctNode" /></xsl:attribute></input>
				<input type="hidden" name="mp"><xsl:attribute name="value"><xsl:value-of select="//queryItems/mp" /></xsl:attribute></input>
				<input type="hidden" name="kpi"><xsl:attribute name="value"><xsl:value-of select="//queryItems/kpi" /></xsl:attribute></input>
				<label for="voteResult">
					<input name="voteResult" id="voteResult1" type="radio" value="A" />
					<img alt="一級棒" title="一級棒">
						<xsl:attribute name="src">xslgip/<xsl:value-of select="myStyle"/>/images/exp1.png</xsl:attribute>
					</img>
				</label>
				<label for="voteResult">
					<input name="voteResult" id="voteResult2" type="radio" value="B" />
					<img alt="我喜歡" title="我喜歡">
						<xsl:attribute name="src">xslgip/<xsl:value-of select="myStyle"/>/images/exp2.png</xsl:attribute>
					</img>
				</label>
				<label for="voteResult">
					<input name="voteResult" id="voteResult3" type="radio" value="C" />
					<img alt="很實用" title="很實用">
						<xsl:attribute name="src">xslgip/<xsl:value-of select="myStyle"/>/images/exp3.png</xsl:attribute>
					</img>
				</label>
				<label for="voteResult">
					<input name="voteResult" id="voteResult4" type="radio" value="D" />
					<img alt="夠新奇" title="夠新奇">
						<xsl:attribute name="src">xslgip/<xsl:value-of select="myStyle"/>/images/exp4.png</xsl:attribute>
					</img>
				</label>
				<label for="voteResult">
					<input name="voteResult" id="voteResult5" type="radio" value="E" />
					<img alt="普普啦" title="普普啦">
						<xsl:attribute name="src">xslgip/<xsl:value-of select="myStyle"/>/images/exp5.png</xsl:attribute>
					</img>
				</label>
				<textarea name="voteDes" wrap="virtual" onclick="ClearDefault()">我有話要說</textarea>
				<input name="SubmitBtn" id="SubmitBtn" type="button" value="送出" onclick="SendVote()" />
			</form>
			<!-- div class="more"><a href ="#">瀏覽更多態度投票結果</a></div -->
		</div>
		<script language="javascript">
			<![CDATA[
			function SendVote() {
				var flag = false;
				if( document.getElementById("voteDes").value == "我有話要說" ) {
					document.getElementById("voteDes").value = "";
				}
				for( var i = 1; i <= 5; i++ ) {
					if( document.getElementById("voteResult" + i).checked ){
						flag = true;
					}
				}
				if( !flag ) {
					alert("請選擇一項態度!");
				}
				else {
					var str = document.getElementById("voteDes").value;
					str = str.replace( new RegExp("<","gm"), "&lt;");
          str = str.replace( new RegExp(">","gm"), "&gt;");
					str = str.replace( new RegExp("'","gm"), "''");
					str = str.replace( new RegExp("--","gm"), "----");
					document.getElementById("AttitudeVoteForm").submit();
				}
			}
			function ClearDefault() {
				if( document.getElementById("voteDes").value == "我有話要說" ) {
					document.getElementById("voteDes").value = "";
				}
			}
			]]>
		</script>
	</xsl:if>	
		
	<!-- 態度投票結果：開始 -->
	<div class="scored">
		<h5>大家覺得這篇文章</h5>
		<div class="total">共有<em><xsl:value-of select="//attributeVote/totalVote" /></em>人投摽</div>
		<ul>
			<li class="score1" title="一級棒">
				<img>
					<xsl:attribute name="src">xslgip/<xsl:value-of select="myStyle"/>/images/null.gif</xsl:attribute>
					<xsl:attribute name="alt">得票數：<xsl:value-of select="//attributeVote/voteA"/></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="//attributeVote/voteA"/></xsl:attribute>
					<xsl:attribute name="height"><xsl:value-of select="//attributeVote/voteAPercent"/></xsl:attribute>				
				</img>
			</li>
			<li class="score2" title="我喜歡">
				<img>
					<xsl:attribute name="src">xslgip/<xsl:value-of select="myStyle"/>/images/null.gif</xsl:attribute>
					<xsl:attribute name="alt">得票數：<xsl:value-of select="//attributeVote/voteB"/></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="//attributeVote/voteB"/></xsl:attribute>
					<xsl:attribute name="height"><xsl:value-of select="//attributeVote/voteBPercent"/></xsl:attribute>				
				</img>
			</li>
			<li class="score3" title="很實用">
				<img>
					<xsl:attribute name="src">xslgip/<xsl:value-of select="myStyle"/>/images/null.gif</xsl:attribute>
					<xsl:attribute name="alt">得票數：<xsl:value-of select="//attributeVote/voteC"/></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="//attributeVote/voteC"/></xsl:attribute>
					<xsl:attribute name="height"><xsl:value-of select="//attributeVote/voteCPercent"/></xsl:attribute>				
				</img>
			</li>
			<li class="score4" title="夠新奇">
				<img>
					<xsl:attribute name="src">xslgip/<xsl:value-of select="myStyle"/>/images/null.gif</xsl:attribute>
					<xsl:attribute name="alt">得票數：<xsl:value-of select="//attributeVote/voteD"/></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="//attributeVote/voteD"/></xsl:attribute>
					<xsl:attribute name="height"><xsl:value-of select="//attributeVote/voteDPercent"/></xsl:attribute>				
				</img>
			</li>
			<li class="score5" title="普普啦">
				<img>
					<xsl:attribute name="src">xslgip/<xsl:value-of select="myStyle"/>/images/null.gif</xsl:attribute>
					<xsl:attribute name="alt">得票數：<xsl:value-of select="//attributeVote/voteE"/></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="//attributeVote/voteE"/></xsl:attribute>
					<xsl:attribute name="height"><xsl:value-of select="//attributeVote/voteEPercent"/></xsl:attribute>				
				</img>
			</li>
		</ul>
		<!-- div class="more"><a href ="#">瀏覽更多態度投票結果</a></div -->
		<h5>看過這篇文章的人說</h5>
		<ol>
			<xsl:for-each select="//attributeVote/article">
				<li>
					<xsl:value-of select="xBody" />
				</li>
			</xsl:for-each>			
		</ol>
	</div>
	<!-- 態度投票結果：結束 -->

</xsl:template>	

</xsl:stylesheet>
