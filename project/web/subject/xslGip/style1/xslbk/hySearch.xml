<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
	<xsl:template match="hpMain" mode="xsSearch">
					<div class="Search">
						<div class="Head">搜尋本站</div>
						<form name="xgils_search" method="GET" action="http://demo.hyweb.com.tw/hysearch/cgi/m_query.exe"
							target="_nwGIP">
							<div id="search">
								<input type="hidden" name="home" value="home" />
								<input type="hidden" name="path" value="/HyWeb/database/pages" />
								<input type="hidden" name="sort_type" value="sort_h" />
								<input type="hidden" name="sort_type" value="hostname" />
								<input type="hidden" name="item_no" value="10" />
								<input type="hidden" name="phonetic" value="0" />
								<input type="hidden" name="fuzzy" value="0" />
								<input type="hidden" name="nature" value="0" />
								<input type="hidden" name="group" value="+" />
								<input type="hidden" name="check_group" value="yes" />
								<input type="hidden" name="template" value="s" />
								<input name="content" type="text" class="inputtext" value="請輸入查詢條件" onfocus="clearText(this)"
									size="14" maxlength="20" />
								<input type="button" name="Submit" value="查詢" class="Button" onkeypress="doSearch();" onclick="doSearch();"/>
							</div> <!--id="search"-->
<script language="JavaScript" type="text/JavaScript">
function clearText(x) {
	if ("請輸入查詢條件"==x.value)	x.value="";
}

function doSearch() {
window.open(encodeURI("http://theabs.hyweb.com.tw/hysearch/cgi/m_query.exe?home=home&amp;path=/HyWeb/database/pages&amp;sort_type=sort_h&amp;sort_field=hostname&amp;item_no=10&amp;phonetic=0&amp;fuzzy=0&amp;nature=0&amp;group=+&amp;check_group=yes&amp;content="+xgils_search.content.value), "_blank");
}
</script>						</form>
					</div>
	</xsl:template>
</xsl:stylesheet>
