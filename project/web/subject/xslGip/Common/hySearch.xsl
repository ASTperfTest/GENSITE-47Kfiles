<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:template match="hpMain" mode="xsSearch">

    <form name="SearchForm" id="SearchForm" method="post" taget="_self" >
      <label for="Keyword" accesskey="S">網站搜尋: </label>					
			<input name="Keyword" id="Keyword" type="text" class="input" value="請輸入查詢條件" onfocus="this.value='';" size="18" maxlength="20" />
			<input name="FromSiteUnit" id="FromSiteUnit" value="1" type="hidden" />
			<input name="FromKnowledgeTank" id="FromKnowledgeTank" value="1" type="hidden" />
			<input name="FromKnowledgeHome" id="FromKnowledgeHome" value="1" type="hidden" />
			<input name="FromTopic" id="FromTopic" value="1" type="hidden" />
      			
		 	<input name="search" id="search" type="button" value="查詢" class="button" onClick="javascript:checkSearchForm()" />		  
		  
		</form>
		
		<script language="javascript">
		
		function checkSearchForm()
		{			
			if( document.SearchForm.Keyword.value.replace(/^[\s　]+/g, "").replace(/[\s　]+$/g, "")== "" ) {
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
		
	
	</xsl:template>
</xsl:stylesheet>
