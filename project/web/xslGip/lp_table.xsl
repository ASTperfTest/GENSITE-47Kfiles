<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
xmlns:user="urn:user-namespace-here" version="1.0">

<xsl:template match="TopicList">

<!--篩選開始-->
<xsl:if test="@xNode='1584xxx'">

<script language="vbs"><![CDATA[
	sub htx_topCat_onChange
 		set xsrc = document.all("htx_vgroup")
		removeOption(xsrc)
		xaddOption xsrc, "不分類", ""

	  set oXML = createObject("Microsoft.XMLDOM")
 		oXML.async = false

	 	xURI = "/ws/lawVerID.asp?lawID=" & reg.htx_topCat.value
 		oXML.load(xURI)

		

	 	set pckItemList = oXML.selectNodes("divList/row")
 		for each pckItem in pckItemList
  			xaddOption xsrc, pckItem.selectSingleNode("mValue").text,pckItem.selectSingleNode("mCode").text
	 	next
 		xsrc.selectedIndex = 0		
	end sub
	
	sub removeOption(xlist)
 		for i=xlist.options.length-1 to 0 step -1
  			xlist.options.remove(i)
	 	next
 		xlist.selectedIndex = -1
	end sub

	sub xaddOption(xlist,name,value)
	 	set xOption = document.createElement("OPTION")
	 	xOption.text=name
	 	xOption.value=value
	 	xlist.add(xOption)
	end sub

]]></script>
 <form method="POST" name="reg" action="http://kmweb.coa.gov.tw/lp.asp?CtNode=1584&amp;CtUnit=310&amp;BaseDSD=41&amp;mp=1">
 
 
 大分類：<select name="htx_topCat" size="1" title="影音分類一" tabindex="">
  <option value="">請選擇</option> 
  <option value="01">頑皮家族</option> 
  <option value="15">因材施教</option> 
  <option value="20">山光水色</option> 
  <option value="35">農妝堰抹</option> 
  <option value="45">綠野仙蹤</option> 
  <option value="55">霽月風光</option> 
  </select>
  
 次分類：<select name="htx_vgroup" size="1" title="不勾為全合適" tabindex="">
  <option value="">請選擇</option> 
  <!--option value="111133">飛禽</option> 
  <option value="111134">千山</option> 
  <option value="111135">村莊</option> 
  <option value="111136">百花</option> 
  <option value="111141">水利</option> 
  <option value="111143">綠園</option> 
  <option value="111144">綠園</option> 
  <option value="111142">水利</option> 
  <option value="111139">萬水</option> 
  <option value="111137">走獸</option> 
  <option value="111138">游兒</option> 
  <option value="111140">林蔭</option--> 
  </select>
  
  <input type="submit" value="篩選" class="cbutton" /> 

  </form>
<!--篩選結束-->
</xsl:if>
	<div class="lp">
		<xsl:value-of disable-output-escaping="yes" select="../HeaderPart" />
		<!--標准文章樣式-->
		<table width="100%" border="0" cellspacing="0" cellpadding="0" summary="問題列表資料表格">
		<xsl:apply-templates select="../TopicTitleList" />		
				<xsl:for-each select="Article">
					<tr>
						<td width="5%"><xsl:value-of select="position()"/></td>
						<xsl:for-each select="ArticleField">
							<xsl:apply-templates select="."/>
						</xsl:for-each>
					</tr>
				</xsl:for-each>
		</table>
		<xsl:value-of disable-output-escaping="yes" select="../FooterPart" />
	</div>
</xsl:template>
    
<xsl:template match="TopicTitleList">    
  	<tr>
		<th scope="col">&amp;nbsp;</th>
		<xsl:for-each select="TopicTitle"> 			
			<th scope="col"><xsl:value-of select="." /></th>
		</xsl:for-each>			
	</tr>
</xsl:template>

<xsl:template match="ArticleField">    
	<xsl:choose>	
		<xsl:when test="xURL">
			<td>
				<a>
					<xsl:attribute name="href"><xsl:value-of disable-output-escaping="yes" select="xURL" /></xsl:attribute>
					<xsl:attribute name="title"><xsl:value-of select="Value" /></xsl:attribute> 
					<xsl:if test="xURL[@newWindow='Y']">
						<xsl:attribute name="target">_blank</xsl:attribute>
					</xsl:if>
					<xsl:value-of select="Value" />       	        		
				</a>			
			</td>
		</xsl:when>
		<xsl:when test="Value=''"><td>&amp;nbsp;</td></xsl:when>
		<xsl:otherwise><td><xsl:value-of select="Value" /></td></xsl:otherwise>
	</xsl:choose>	
</xsl:template>


	
	
	<!-- 資料大類 -->
<xsl:template match="CatList">
    <div class="category">
<span class="cateName">主分類</span>
<ul>
<li><a><xsl:attribute name="href">lp.asp?<xsl:value-of select="../xqURL" /></xsl:attribute>
        <xsl:if test="not(contains(//qURL, 'xq_xCat='))"><b>All</b></xsl:if>
        <xsl:if test="(contains(//qURL, 'xq_xCat='))">All</xsl:if>
        </a></li>
    		<xsl:for-each select="CatItem">
<li>
    			<a>
    				<xsl:attribute name="href">lp.asp?<xsl:value-of select="../../xqURL" /><xsl:value-of select="xqCondition" /></xsl:attribute>
    				<xsl:if test="contains(concat(//qURL, ''), concat(xqCondition, ''))"><b>
            <xsl:value-of select="CatName" /></b>
            </xsl:if>
            	<xsl:if test="not(contains(concat(//qURL, ''), concat(xqCondition, '')))">
            <xsl:value-of select="CatName" />
            </xsl:if>
    			</a>
</li>
    		</xsl:for-each>
</ul>
    </div>
    		<xsl:apply-templates select="//CatListRelated" />
</xsl:template>
<!-- 資料大類 -->

<xsl:template match="CatListRelated">
<div class="category">
<span class="cateName">次分類</span>
<ul>
  <xsl:for-each select="row">
<li>
      <a>
        <xsl:attribute name="href">lp.asp?<xsl:if test="contains(//qURL, 'htx_vgroup=')"><xsl:value-of select="//xqURLDeleteAnd1"/>&amp;htx_vgroup=<xsl:value-of select="mCode"/>
        </xsl:if>
          
          <xsl:if test="not(contains(//qURL, 'htx_vgroup='))"><xsl:value-of select="//qURL"/>&amp;htx_vgroup=<xsl:value-of select="mCode"/>
          </xsl:if>
        </xsl:attribute>
        <xsl:if test="contains(concat(//qURL, ''), concat(mCode, ''))"><b>
        <xsl:value-of select="mValue" /></b>
         </xsl:if>
         
           <xsl:if test="not(contains(concat(//qURL, ''), concat(mCode, '')))">
        <xsl:value-of select="mValue" />
         </xsl:if>
         
      </a>
</li>
    </xsl:for-each>
</ul>
</div> 
</xsl:template>


<!--路徑連結 Start-->
<xsl:template match="hpMain" mode="xsPath">
  <div class="path" style="float:left; padding-top: 10px">目前位置：</div>
  <div style="float:left;width:80%">
    <ul id="path_menu">
      <li>
        <a title="首頁">
          <xsl:attribute name="href">
            mp.asp?mp=<xsl:value-of select="//mp" />
          </xsl:attribute>
          首頁
        </a>
      </li>
      <xsl:for-each select="//xPath/xPathNode">
        <li style="top: 10px;">></li>
        <li>
          <a>
            <xsl:attribute name="href">
              np.asp?ctNode=<xsl:value-of select="@xNode" />&amp;mp=<xsl:value-of select="//mp" />
            </xsl:attribute>
            <!--<xsl:attribute name="title">
            <xsl:value-of select="@Title" />
          </xsl:attribute>-->
            <xsl:value-of select="@Title" />
          </a>
        </li>
      </xsl:for-each>
      <xsl:for-each select="//CatList/mediapath">
        <xsl:if test="mcode!=''">
          <li style="top: 10px;">></li>
          <xsl:choose>
            <xsl:when test="murl!=''">
              <li>
                <a>
                  <xsl:attribute name="href">
                    <xsl:value-of select="murl" />&amp;mp=<xsl:value-of select="//mp" />
                  </xsl:attribute>
                  <xsl:attribute name="title">
                    <xsl:value-of select="mvalue" />
                  </xsl:attribute>
                  <xsl:value-of select="mvalue" />
                </a>
              </li>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="mvalue" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:if>
      </xsl:for-each>
    </ul>
  </div>  
	<!--<div class="path">目前位置：
		<a title="首頁">
			<xsl:attribute name="href">mp.asp?mp=<xsl:value-of select="//mp" /></xsl:attribute>
			首頁
		</a>
		<xsl:for-each select="//xPath/xPathNode">
			>
			<a>
				<xsl:attribute name="href">np.asp?ctNode=<xsl:value-of select="@xNode" />&amp;mp=<xsl:value-of select="//mp" /></xsl:attribute>
				<xsl:attribute name="title"><xsl:value-of select="@Title" /></xsl:attribute>
				<xsl:value-of select="@Title" />
			</a>
		</xsl:for-each>
		
        --><!--路徑套用主分類 START--><!--
        <xsl:for-each select="//CatList/CatItem">
        <xsl:if test="contains(concat(//qURL, ''), concat(xqCondition, ''))">> </xsl:if>
       
        <xsl:if test="contains(concat(//qURL, ''), concat(xqCondition, ''))">
         <a>
        <xsl:attribute name="href">lp.asp?<xsl:value-of select="../../xqURL" /><xsl:value-of select="xqCondition" /></xsl:attribute>
        <xsl:attribute name="title"><xsl:value-of select="CatName" /></xsl:attribute>
        <xsl:value-of select="CatName" />
        </a>
        </xsl:if>
        
        </xsl:for-each>
        --><!--路徑套用主分類 END--><!--
        --><!--路徑套用次分類 START--><!--
        <xsl:for-each select="//CatListRelated/row">
        <xsl:if test="contains(concat(//qURL, ''), concat(mCode, ''))">> </xsl:if>
        
        <xsl:if test="contains(concat(//qURL, ''), concat(mCode, ''))">
        <a>
          <xsl:attribute name="href">lp.asp?<xsl:if test="contains(//qURL, 'htx_vgroup=')"><xsl:value-of select="//xqURLDeleteAnd1"/>&amp;htx_vgroup=<xsl:value-of select="mCode"/></xsl:if>
          </xsl:attribute> 
          <xsl:attribute name="title"><xsl:value-of select="mValue" /></xsl:attribute>
        <xsl:value-of select="mValue" />
        </a>
        </xsl:if>
        
        </xsl:for-each>
        --><!--路徑套用次分類 END--><!--
	  </div>-->
</xsl:template>
<!--路徑連結 End-->
</xsl:stylesheet>

