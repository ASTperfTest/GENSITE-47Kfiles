<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:saxon="http://icl.com/saxon/"
    extension-element-prefixes="saxon" version="1.0">
    <xsl:include href="accesskey.xsl" />
    <!--xsl:include href="member.xsl" /-->
    <xsl:variable name="title">農委會鳳梨知識主題館</xsl:variable>
    <!-- 網站標頭／開始 -->
    <xsl:template match="hpMain" mode="topinfo">
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>
            <xsl:value-of select="$title" />
        </title>
    </xsl:template>
    <!-- 網站標頭／結束 -->
    <!-- 網站標題／開始 -->
    <xsl:template match="hpMain" mode="webinfo">
        <xsl:apply-templates select="." mode="u" />
        <div class="WebTitle">
            <img>
                <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/webtitle01.jpg</xsl:attribute>
                <xsl:attribute name="alt">
                    <xsl:value-of select="$title" />
                </xsl:attribute>
            </img>
        </div>
    </xsl:template>
    <!-- 網站標題／結束 -->
    <!-- 導覽列／開始 -->
    <xsl:template match="hpMain" mode="info">
     <link rel="stylesheet" type="text/css">
		  	<xsl:attribute name="href">xslgip/<xsl:value-of select="//MpStyle" />/css/login.css</xsl:attribute>
         </link >
		 <link rel="stylesheet" type="text/css">
		                       <xsl:attribute name="href">/css/pedia.css</xsl:attribute>
	                </link>
		  <script src="/js/pedia.js" language="javascript" />
					<script src="/js/jquery.js" language="javascript" />
			<!--div id="login" ><xsl:apply-templates select="." mode="xsMember" /></div-->
        <ul class="TopLink">
			<li><xsl:if test="//login/status='true'">
				<a><xsl:attribute name="href">/sp.asp?xdURL=Coamember/member_Modify.asp&amp;mp=1</xsl:attribute><xsl:value-of select="//login/memAccount" /></a><a><xsl:attribute name="href">logout.asp</xsl:attribute>登出</a>
			</xsl:if>
			<xsl:if test="//login/status='false'">
				<a><xsl:attribute name="href">login.asp</xsl:attribute>登入</a>
			</xsl:if>
			</li>
            <li>
                <a title="首頁"><xsl:attribute name="href">dp.asp?mp=<xsl:value-of select="//mp" /></xsl:attribute>首頁</a>
            </li>
            <li>
                <!--<a title="聯絡我們"><xsl:apply-templates select="." mode="ContactUs" />聯絡我們</a>-->
                <!--a title="聯絡我們" href="mailto:liucc@hyweb.com.tw">聯絡我們</a-->
            </li>
            <li>
                <a title="導覽"><xsl:attribute name="href">sitemap.asp?mp=<xsl:value-of select="//Mp" /></xsl:attribute>導覽</a>
                <!--a title="網站導覽" href="#">網站導覽</a-->
            </li>
			<li>
            <a title="主題列表" href="/SubjectList.aspx">
              主題列表
            </a>
          </li>	
			<li>
            <a title="入口網" href="/">
              入口網
            </a>
          </li>
        </ul>
        	<!--script src="/nocopy.js" language="javascript"></script-->
        	  <!-- 搜尋／開始 -->
          <div class="search">
          <form name="SearchForm" method="post" taget="_blank" >
          <label  for="Keyword" accesskey="S">網站搜尋：</label>
          <input id="content" name="Keyword" type="text" class="input" onfocus="this.value='';" value="請輸入關鍵字" size="18" />
          <input name="FromSiteUnit" value="1" type="hidden" />
          <input name="FromKnowledgeTank" value="1" type="hidden" />
          <input name="FromKnowledgeHome" value="1" type="hidden" />
          <input name="FromTopic" value="1" type="hidden" />
          <input name="search" type="button" value="查詢" class="button"  onClick="javascript:checkSearchForm()" /> 
          </form>
          </div>
              <script language="javascript">
              
              function checkSearchForm()
              {			
              if( document.SearchForm.Keyword.value == "" ) {
              alert('請輸入查詢值');
              event.returnValue = false;
              }
              else {
              
              document.SearchForm.FromSiteUnit.value = document.SearchForm.FromSiteUnit.value.replace(new RegExp("<xsl:text disable-output-escaping="yes">'</xsl:text>","gm"), "''");
              document.SearchForm.FromSiteUnit.value = document.SearchForm.FromSiteUnit.value.replace(new RegExp("<xsl:text disable-output-escaping="yes">></xsl:text>","gm"), "&gt;");
              document.SearchForm.FromSiteUnit.value = document.SearchForm.FromSiteUnit.value.replace(new RegExp("<xsl:text disable-output-escaping="yes"><![CDATA[<]]></xsl:text>","gm"), "&lt;");
              
              document.SearchForm.FromKnowledgeTank.value = document.SearchForm.FromKnowledgeTank.value.replace(new RegExp("<xsl:text disable-output-escaping="yes">'</xsl:text>","gm"), "''");
              document.SearchForm.FromKnowledgeTank.value = document.SearchForm.FromKnowledgeTank.value.replace(new RegExp("<xsl:text disable-output-escaping="yes">></xsl:text>","gm"), "&gt;");
              document.SearchForm.FromKnowledgeTank.value = document.SearchForm.FromKnowledgeTank.value.replace(new RegExp("<xsl:text disable-output-escaping="yes"><![CDATA[<]]></xsl:text>","gm"), "&lt;");
              
              document.SearchForm.FromKnowledgeHome.value = document.SearchForm.FromKnowledgeHome.value.replace(new RegExp("<xsl:text disable-output-escaping="yes">'</xsl:text>","gm"), "''");
              document.SearchForm.FromKnowledgeHome.value = document.SearchForm.FromKnowledgeHome.value.replace(new RegExp("<xsl:text disable-output-escaping="yes">></xsl:text>","gm"), "&gt;");
              document.SearchForm.FromKnowledgeHome.value = document.SearchForm.FromKnowledgeHome.value.replace(new RegExp("<xsl:text disable-output-escaping="yes"><![CDATA[<]]></xsl:text>","gm"), "&lt;");
              
              document.SearchForm.FromTopic.value = document.SearchForm.FromTopic.value.replace(new RegExp("<xsl:text disable-output-escaping="yes">'</xsl:text>","gm"), "''");
              document.SearchForm.FromTopic.value = document.SearchForm.FromTopic.value.replace(new RegExp("<xsl:text disable-output-escaping="yes">></xsl:text>","gm"), "&gt;");
              document.SearchForm.FromTopic.value = document.SearchForm.FromTopic.value.replace(new RegExp("<xsl:text disable-output-escaping="yes"><![CDATA[<]]></xsl:text>","gm"), "&lt;");
              
              var str = document.SearchForm.Keyword.value;
              str = str.replace(new RegExp("<xsl:text disable-output-escaping="yes">'</xsl:text>","gm"), "''");
              str = str.replace(new RegExp("<xsl:text disable-output-escaping="yes">></xsl:text>","gm"), "&gt;");
              str = str.replace(new RegExp("<xsl:text disable-output-escaping="yes"><![CDATA[<]]></xsl:text>","gm"), "&lt;");
              document.SearchForm.Keyword.value = str;
              document.SearchForm.action = "/kp.asp?xdURL=Search/SearchResultList.asp<xsl:text disable-output-escaping="yes">&amp;</xsl:text>mp=1";
              document.SearchForm.submit();				
              }		
              }
              </script>
          <!-- 搜尋／結束 -->
    </xsl:template>
    <!-- 導覽列／結束 -->
    <!-- 站內查詢／開始 -->
    <xsl:template match="hpMain" mode="Search">
        <form name="hysearchform" id="form1" method="post" action="" class="Qsearch">
		<xsl:apply-templates select="." mode="s" />
		<label>站內查詢
			<input name="content" type="text" size="12" class="Text" value="請輸入關鍵字" onfocus="JSCRIPT: this.value='';" />
		</label>&amp;nbsp;
		<input type="submit" name="query" value="查詢" class="Button" onclick="JSCRIPT: return hysearch(this.form);"
                onkeypress="JSCRIPT: return hysearch(this.form);" title="查詢" />
		<a>
			<xsl:attribute name="href">http://search.xxx.gov.tw/hysearch/dindex.htm</xsl:attribute>
			<xsl:attribute name="title">進階查詢[另開視窗]</xsl:attribute>
			<xsl:attribute name="target">Dhysearch</xsl:attribute>進階查詢		
		</a>			
		<a href="#" title="說明">說明</a>
	</form>
        <script language="javascript">
		function hysearch(form) {
		var op = "%2B"
		window.open(encodeURI("http://search.xxx.gov.tw/hysearch/cgi/m_query.exe?content=" + form.content.value + "&amp;path=/HyWeb5.0.1/database/pages&amp;item_no=10&amp;check_group=yes&amp;group=HH"),'_blank');
		}
		</script>
        <noscript>本網頁使用SCRIPT編碼方式轉換您輸入的查詢關鍵字，如果您的瀏覽器不支援SCRIPT，請直接前往<a href="http://search.xxx.gov.tw/hysearch/" target="hysearch">搜尋引擎首頁</a>進行網站搜尋</noscript>
    </xsl:template>
    <!-- 站內查詢／結束 -->
    <!-- 訂閱電子報／開始 -->
    <xsl:template match="hpMain" mode="xsEPaper">
        <div class="Epaper">
            <h3>訂閱電子報</h3>
            <div class="Body">
                <form method="post" action="member/epaper_act.asp?CtRootID=#">
				<input name="textfield" type="text" size="20" class="Text" value="e-mail address" onfocus="JSCRIPT: this.value='';" /><br />
				<input name="submit1" type="submit" value="訂閱" class="Button" title="確定訂閱電子報" />&amp;nbsp;
				<input name="submit2" type="submit" value="取消訂閱" class="Button" title="取消訂閱電子報" />
			</form>
            </div>
            <div class="Foot"></div>
        </div>
    </xsl:template>
    <!-- 訂閱電子報／結束 -->
</xsl:stylesheet>
