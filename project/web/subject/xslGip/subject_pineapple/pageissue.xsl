<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
    xmlns:user="urn:user-namespace-here" version="1.0">
    <!-- 中文分頁／開始-->
    <xsl:template match="hpMain" mode="pageissueC">
        <xsl:variable name="PerPageSize">
            <xsl:value-of select="PerPageSize" />
        </xsl:variable>
        <!--xsl:if test="totRec>$PerPageSize"-->
        <div class="Page"> 
			共 <span class="Number">
                <xsl:value-of select="totPage" />
            </span> 頁， <span class="Number">
                <xsl:value-of select="totRec" />
            </span> 筆資料
			<xsl:if test="nowPage > 1">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="上一頁">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/bullet13.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>			
			 &amp;nbsp;到第
			<select name="pickPage" id="pickPage" class="inputtext" onChange="pageChange(this.value)">
                <script>
					var totPage = <xsl:value-of select="totPage" />;
					var qURL = "<xsl:value-of select="qURL" />";
					var nowPage = "<xsl:value-of select="nowPage" />";
					var PerPageSize = "<xsl:value-of select="PerPageSize" />";
					<xsl:text disable-output-escaping="yes" >
					<![CDATA[
					for (xi=1;xi<=totPage;xi++)
					if (xi==nowPage)
					document.write("<option value=" +xi + " selected='selected'>" + xi + "</option>");
					else
					document.write("<option value=" +xi + ">" + xi + "</option>");
					
					function pageChange(nPage) {
					//alert(nPage);
					//goPage(pickPage.value);
					goPage(nPage);
					}
					function goPage(nPage) {
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + PerPageSize
					}
					function perPageChange(pagesize) {
					//document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nowPage + "&pagesize=" + pagesize
					}
					]]>
					</xsl:text>
				</script>
            </select>
			頁
			<xsl:if test="number(nowPage) &lt; number(totPage)">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="下一頁">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/icon_men01_1.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>
			&amp;nbsp; 每頁顯示
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
                <option selected="selected" value="0">請選擇</option>
                <option value="10">10</option>
                <option value="20">20</option>
                <option value="30">30</option>
            </select> 筆
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
			<noscript>
				<br />每頁<xsl:value-of select="PerPageSize" />筆,目前在第<xsl:value-of select="nowPage" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br /><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
        <!--/xsl:if-->
    </xsl:template>
    <!-- 中文分頁／結束-->
    <!-- 英文分頁／開始-->
    <xsl:template match="hpMain" mode="pageissueE">
        <xsl:variable name="PerPageSize">
            <xsl:value-of select="PerPageSize" />
        </xsl:variable>
        <!--xsl:if test="totRec>$PerPageSize"-->
        <div class="Page"> 
			Total <span class="Number">
                <xsl:value-of select="totPage" />
            </span> Page(s)， <span class="Number">
                <xsl:value-of select="totRec" />
            </span>record(s)，
			<xsl:if test="nowPage > 1">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="Previous">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_prev.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>			
			 &amp;nbsp;go to page
			<select name="pickPage" id="pickPage" class="inputtext" onChange="pageChange(this.value)">
                <script>
					var totPage = <xsl:value-of select="totPage" />;
					var qURL = "<xsl:value-of select="qURL" />";
					var nowPage = "<xsl:value-of select="nowPage" />";
					var PerPageSize = "<xsl:value-of select="PerPageSize" />";
					<![CDATA[
					for (xi=1;xi<=totPage;xi++)
					if (xi==nowPage)
					document.write("<option value=" +xi + " selected='selected'>" + xi + "</option>");
					else
					document.write("<option value=" +xi + ">" + xi + "</option>");
					
					function pageChange(nPage) {
					//alert(nPage);
					//goPage(pickPage.value);
					goPage(nPage);
					}
					function goPage(nPage) {
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + PerPageSize
					}
					function perPageChange(pagesize) {
					//document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nowPage + "&pagesize=" + pagesize
					}
					]]>
				</script>
            </select>
			

			<xsl:if test="number(nowPage) &lt; number(totPage)">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="Next">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_next.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>
			&amp;nbsp; Display
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
                <option selected="selected">
					<xsl:attribute name="value">0<xsl:value-of select="PerPageSize" /></xsl:attribute>
					----
				</option>
                <option value="15">15</option>
                <option value="30">30</option>
                <option value="50">50</option>
            </select> record(s) per page
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
			<noscript>
				<br />每頁<xsl:value-of select="PerPageSize" />筆,目前在第<xsl:value-of select="nowPage" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br /><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
        <!--/xsl:if-->
    </xsl:template>
    <!-- 英文分頁／結束-->
    <!-- 日文分頁／開始-->
    <xsl:template match="hpMain" mode="pageissueJ">
        <xsl:variable name="PerPageSize">
            <xsl:value-of select="PerPageSize" />
        </xsl:variable>
        <!--xsl:if test="totRec>$PerPageSize"-->
        <div class="Page"> 
			計 <span class="Number">
                <xsl:value-of select="totPage" />
            </span> ???， 資料件? <span class="Number">
                <xsl:value-of select="totRec" />
            </span> 件
			<xsl:if test="nowPage > 1">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="前頁">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_prev.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>			
			 &amp;nbsp;到第
			<select name="pickPage" id="pickPage" class="inputtext" onChange="pageChange(this.value)">
                <script>
					var totPage = <xsl:value-of select="totPage" />;
					var qURL = "<xsl:value-of select="qURL" />";
					var nowPage = "<xsl:value-of select="nowPage" />";
					var PerPageSize = "<xsl:value-of select="PerPageSize" />";
					<![CDATA[
					for (xi=1;xi<=totPage;xi++)
					if (xi==nowPage)
					document.write("<option value=" +xi + " selected='selected'>" + xi + "</option>");
					else
					document.write("<option value=" +xi + ">" + xi + "</option>");
					
					function pageChange(nPage) {
					//alert(nPage);
					//goPage(pickPage.value);
					goPage(nPage);
					}
					function goPage(nPage) {
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + PerPageSize
					}
					function perPageChange(pagesize) {
					//document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nowPage + "&pagesize=" + pagesize
					}
					]]>
				</script>
            </select>
			頁

			<xsl:if test="number(nowPage) &lt; number(totPage)">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="次頁">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_next.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>
			&amp;nbsp; 各頁
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
                <option selected="selected">
					<xsl:attribute name="value">0<xsl:value-of select="PerPageSize" /></xsl:attribute>
					----
				</option>
                <option value="15">15</option>
                <option value="30">30</option>
                <option value="50">50</option>
            </select> 件??表示
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
			<noscript>
				<br />各頁<xsl:value-of select="PerPageSize" />件??表示,目前在第<xsl:value-of select="nowPage" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br /><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
        <!--/xsl:if-->
    </xsl:template>
    <!-- 日文分頁／結束-->
    <!-- 韓文分頁／開始-->
    <xsl:template match="hpMain" mode="pageissueK">
        <xsl:variable name="PerPageSize">
            <xsl:value-of select="PerPageSize" />
        </xsl:variable>
        <!--xsl:if test="totRec>$PerPageSize"-->
        <div class="Page"> 
			? <span class="Number">
                <xsl:value-of select="totPage" />
            </span> ???， <span class="Number">
                <xsl:value-of select="totRec" />
            </span> ???
			<xsl:if test="nowPage > 1">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="??">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_prev.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>			
			 &amp;nbsp;?
			<select name="pickPage" id="pickPage" class="inputtext" onChange="pageChange(this.value)">
                <script>
					var totPage = <xsl:value-of select="totPage" />;
					var qURL = "<xsl:value-of select="qURL" />";
					var nowPage = "<xsl:value-of select="nowPage" />";
					var PerPageSize = "<xsl:value-of select="PerPageSize" />";
					<![CDATA[
					for (xi=1;xi<=totPage;xi++)
					if (xi==nowPage)
					document.write("<option value=" +xi + " selected='selected'>" + xi + "</option>");
					else
					document.write("<option value=" +xi + ">" + xi + "</option>");
					
					function pageChange(nPage) {
					//alert(nPage);
					//goPage(pickPage.value);
					goPage(nPage);
					}
					function goPage(nPage) {
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + PerPageSize
					}
					function perPageChange(pagesize) {
					//document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nowPage + "&pagesize=" + pagesize
					}
					]]>
				</script>
            </select>
			???

			<xsl:if test="number(nowPage) &lt; number(totPage)">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="??">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_next.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>
			&amp;nbsp; ????
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
                <option selected="selected">
					<xsl:attribute name="value">0<xsl:value-of select="PerPageSize" /></xsl:attribute>
					----
				</option>
                <option value="15">15</option>
                <option value="30">30</option>
                <option value="50">50</option>
            </select> ???
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
			<noscript>
				<br />????<xsl:value-of select="PerPageSize" />???,目前在第<xsl:value-of select="nowPage" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br /><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
        <!--/xsl:if-->
    </xsl:template>
    <!-- 韓文分頁／結束-->
    <!-- 西文分頁／開始-->
    <xsl:template match="hpMain" mode="pageissueS">
        <xsl:variable name="PerPageSize">
            <xsl:value-of select="PerPageSize" />
        </xsl:variable>
        <!--xsl:if test="totRec>$PerPageSize"-->
        <div class="Page"> 
			total <span class="Number">
                <xsl:value-of select="totPage" />
            </span> paginas， <span class="Number">
                <xsl:value-of select="totRec" />
            </span> datos
			<xsl:if test="nowPage > 1">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="pagina anterior">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_prev.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>			
			 &amp;nbsp;a la pagina
			<select name="pickPage" id="pickPage" class="inputtext" onChange="pageChange(this.value)">
                <script>
					var totPage = <xsl:value-of select="totPage" />;
					var qURL = "<xsl:value-of select="qURL" />";
					var nowPage = "<xsl:value-of select="nowPage" />";
					var PerPageSize = "<xsl:value-of select="PerPageSize" />";
					<![CDATA[
					for (xi=1;xi<=totPage;xi++)
					if (xi==nowPage)
					document.write("<option value=" +xi + " selected='selected'>" + xi + "</option>");
					else
					document.write("<option value=" +xi + ">" + xi + "</option>");
					
					function pageChange(nPage) {
					//alert(nPage);
					//goPage(pickPage.value);
					goPage(nPage);
					}
					function goPage(nPage) {
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + PerPageSize
					}
					function perPageChange(pagesize) {
					//document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nowPage + "&pagesize=" + pagesize
					}
					]]>
				</script>
            </select>
			

			<xsl:if test="number(nowPage) &lt; number(totPage)">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="Pagina siguiente">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_next.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>
			&amp;nbsp; cada paginas contiene
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
                <option selected="selected">
					<xsl:attribute name="value">0<xsl:value-of select="PerPageSize" /></xsl:attribute>
					----
				</option>
                <option value="15">15</option>
                <option value="30">30</option>
                <option value="50">50</option>
            </select> datos
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
			<noscript>
				<br />cada paginas contiene<xsl:value-of select="PerPageSize" />datos,目前在第<xsl:value-of select="nowPage" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br /><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
        <!--/xsl:if-->
    </xsl:template>
    <!-- 西文分頁／結束-->
    <!-- 俄文分頁／開始-->
    <xsl:template match="hpMain" mode="pageissueR">
        <xsl:variable name="PerPageSize">
            <xsl:value-of select="PerPageSize" />
        </xsl:variable>
        <!--xsl:if test="totRec>$PerPageSize"-->
        <div class="Page"> 
			????? <span class="Number">
                <xsl:value-of select="totPage" />
            </span> ????????， <span class="Number">
                <xsl:value-of select="totRec" />
            </span> ??????，
			<xsl:if test="nowPage > 1">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="??????????">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_prev.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>			
			 &amp;nbsp;? ????????
			<select name="pickPage" id="pickPage" class="inputtext" onChange="pageChange(this.value)">
                <script>
					var totPage = <xsl:value-of select="totPage" />;
					var qURL = "<xsl:value-of select="qURL" />";
					var nowPage = "<xsl:value-of select="nowPage" />";
					var PerPageSize = "<xsl:value-of select="PerPageSize" />";
					<![CDATA[
					for (xi=1;xi<=totPage;xi++)
					if (xi==nowPage)
					document.write("<option value=" +xi + " selected='selected'>" + xi + "</option>");
					else
					document.write("<option value=" +xi + ">" + xi + "</option>");
					
					function pageChange(nPage) {
					//alert(nPage);
					//goPage(pickPage.value);
					goPage(nPage);
					}
					function goPage(nPage) {
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + PerPageSize
					}
					function perPageChange(pagesize) {
					//document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nowPage + "&pagesize=" + pagesize
					}
					]]>
				</script>
            </select>
			

			<xsl:if test="number(nowPage) &lt; number(totPage)">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="?????????">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_next.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>
			&amp;nbsp; ?? ?????? ????????
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
                <option selected="selected">
					<xsl:attribute name="value">0<xsl:value-of select="PerPageSize" /></xsl:attribute>
					----
				</option>
                <option value="15">15</option>
                <option value="30">30</option>
                <option value="50">50</option>
            </select> ??????
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
			<noscript>
				<br />?? ?????? ????????<xsl:value-of select="PerPageSize" />??????,目前在第<xsl:value-of select="nowPage" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br /><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
        <!--/xsl:if-->
    </xsl:template>
    <!-- 俄文分頁／結束-->
    <!-- 捷文分頁／開始-->
    <xsl:template match="hpMain" mode="pageissueCR">
        <xsl:variable name="PerPageSize">
            <xsl:value-of select="PerPageSize" />
        </xsl:variable>
        <!--xsl:if test="totRec>$PerPageSize"-->
        <div class="Page"> 
			<span class="Number">
                <xsl:value-of select="totPage" />
            </span> strany， <span class="Number">
                <xsl:value-of select="totRec" />
            </span> polo?ek celkem
			<xsl:if test="nowPage > 1">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="zp?t">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_prev.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>			
			 &amp;nbsp;zp?t na stranu
			<select name="pickPage" id="pickPage" class="inputtext" onChange="pageChange(this.value)">
                <script>
					var totPage = <xsl:value-of select="totPage" />;
					var qURL = "<xsl:value-of select="qURL" />";
					var nowPage = "<xsl:value-of select="nowPage" />";
					var PerPageSize = "<xsl:value-of select="PerPageSize" />";
					<![CDATA[
					for (xi=1;xi<=totPage;xi++)
					if (xi==nowPage)
					document.write("<option value=" +xi + " selected='selected'>" + xi + "</option>");
					else
					document.write("<option value=" +xi + ">" + xi + "</option>");
					
					function pageChange(nPage) {
					//alert(nPage);
					//goPage(pickPage.value);
					goPage(nPage);
					}
					function goPage(nPage) {
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + PerPageSize
					}
					function perPageChange(pagesize) {
					//document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nowPage + "&pagesize=" + pagesize
					}
					]]>
				</script>
            </select>
			

			<xsl:if test="number(nowPage) &lt; number(totPage)">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="vp?ed">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_next.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>
			&amp;nbsp; 
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
                <option selected="selected">
					<xsl:attribute name="value">0<xsl:value-of select="PerPageSize" /></xsl:attribute>
					----
				</option>
                <option value="15">15</option>
                <option value="30">30</option>
                <option value="50">50</option>
            </select> polo?ek na stranu
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
			<noscript>
				<br /><xsl:value-of select="PerPageSize" />polo?ek na stranu,目前在第<xsl:value-of select="nowPage" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br /><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
        <!--/xsl:if-->
    </xsl:template>
    <!-- 捷文分頁／結束-->
    <!-- 荷文分頁／開始-->
    <xsl:template match="hpMain" mode="pageissueN">
        <xsl:variable name="PerPageSize">
            <xsl:value-of select="PerPageSize" />
        </xsl:variable>
        <!--xsl:if test="totRec>$PerPageSize"-->
        <div class="Page"> 
			total <span class="Number">
                <xsl:value-of select="totPage" />
            </span>  pagina’s， <span class="Number">
                <xsl:value-of select="totRec" />
            </span> onderwerpen
			<xsl:if test="nowPage > 1">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="Vorig onderwerp">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_prev.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>			
			 &amp;nbsp;naar pagina
			<select name="pickPage" id="pickPage" class="inputtext" onChange="pageChange(this.value)">
                <script>
					var totPage = <xsl:value-of select="totPage" />;
					var qURL = "<xsl:value-of select="qURL" />";
					var nowPage = "<xsl:value-of select="nowPage" />";
					var PerPageSize = "<xsl:value-of select="PerPageSize" />";
					<![CDATA[
					for (xi=1;xi<=totPage;xi++)
					if (xi==nowPage)
					document.write("<option value=" +xi + " selected='selected'>" + xi + "</option>");
					else
					document.write("<option value=" +xi + ">" + xi + "</option>");
					
					function pageChange(nPage) {
					//alert(nPage);
					//goPage(pickPage.value);
					goPage(nPage);
					}
					function goPage(nPage) {
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + PerPageSize
					}
					function perPageChange(pagesize) {
					//document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nowPage + "&pagesize=" + pagesize
					}
					]]>
				</script>
            </select>
			

			<xsl:if test="number(nowPage) &lt; number(totPage)">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="Volgend onderwerp">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_next.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>
			&amp;nbsp;  
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
                <option selected="selected">
					<xsl:attribute name="value">0<xsl:value-of select="PerPageSize" /></xsl:attribute>
					----
				</option>
                <option value="15">15</option>
                <option value="30">30</option>
                <option value="50">50</option>
            </select> onderwerpen getoond per pagina
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
			<noscript>
				<br /><xsl:value-of select="PerPageSize" />onderwerpen getoond per pagina,目前在第<xsl:value-of select="nowPage" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br /><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
        <!--/xsl:if-->
    </xsl:template>
    <!-- 荷文分頁／結束-->
    <!-- 法文分頁／開始-->
    <xsl:template match="hpMain" mode="pageissueF">
        <xsl:variable name="PerPageSize">
            <xsl:value-of select="PerPageSize" />
        </xsl:variable>
        <!--xsl:if test="totRec>$PerPageSize"-->
        <div class="Page"> 
			<span class="Number">
                <xsl:value-of select="totPage" />
            </span> pages au total， <span class="Number">
                <xsl:value-of select="totRec" />
            </span> informations，
			<xsl:if test="nowPage > 1">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="precedent">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_prev.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>			
			 &amp;nbsp;page
			<select name="pickPage" id="pickPage" class="inputtext" onChange="pageChange(this.value)">
                <script>
					var totPage = <xsl:value-of select="totPage" />;
					var qURL = "<xsl:value-of select="qURL" />";
					var nowPage = "<xsl:value-of select="nowPage" />";
					var PerPageSize = "<xsl:value-of select="PerPageSize" />";
					<![CDATA[
					for (xi=1;xi<=totPage;xi++)
					if (xi==nowPage)
					document.write("<option value=" +xi + " selected='selected'>" + xi + "</option>");
					else
					document.write("<option value=" +xi + ">" + xi + "</option>");
					
					function pageChange(nPage) {
					//alert(nPage);
					//goPage(pickPage.value);
					goPage(nPage);
					}
					function goPage(nPage) {
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + PerPageSize
					}
					function perPageChange(pagesize) {
					//document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage.value
					document.location.href= "lp.asp?" + qURL + "&nowPage=" + nowPage + "&pagesize=" + pagesize
					}
					]]>
				</script>
            </select>
			

			<xsl:if test="number(nowPage) &lt; number(totPage)">
                <a>
                    <xsl:attribute name="href">
						lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
					</xsl:attribute>
                    <img alt="suivant">
                        <xsl:attribute name="src">/xslGip/<xsl:apply-templates select="//MpStyle" />/images/pager_next.gif</xsl:attribute>
                    </img>
                </a>
            </xsl:if>
			&amp;nbsp; 
			<select name="perPage" class="inputtext" onChange="perPageChange(this.value)">
                <option selected="selected">
					<xsl:attribute name="value">0<xsl:value-of select="PerPageSize" /></xsl:attribute>
					----
				</option>
                <option value="15">15</option>
                <option value="30">30</option>
                <option value="50">50</option>
            </select> informations par page
		
			<script>
				pickPage.value = <xsl:value-of select="nowPage" />;
				perPage.value = <xsl:value-of select="PerPageSize" />;
			</script>
			<noscript>
				<br />每頁<xsl:value-of select="PerPageSize" />筆,目前在第<xsl:value-of select="nowPage" />頁;
				<xsl:value-of disable-output-escaping="yes" select="noScriptStr2" />
				<br /><xsl:value-of disable-output-escaping="yes" select="noScriptStr" />
			</noscript>
		</div>
        <!--/xsl:if-->
    </xsl:template>
    <!-- 法文分頁／結束-->
</xsl:stylesheet>
