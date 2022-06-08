<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">


    <!-- 主資料區-->
    <xsl:template match="TopicList">
        <xsl:if test="Article[@addtr='0']">
            <tr>
                <xsl:for-each select="Article[@addtr='0']">

                    <xsl:apply-templates select="." />

                </xsl:for-each>
            </tr>
        </xsl:if>

        <xsl:if test="Article[@addtr='1']">
            <tr>
                <xsl:for-each select="Article[@addtr='1']">

                    <xsl:apply-templates select="." />

                </xsl:for-each>
            </tr>
        </xsl:if>

        <xsl:if test="Article[@addtr='2']">
            <tr>
                <xsl:for-each select="Article[@addtr='2']">

                    <xsl:apply-templates select="." />

                </xsl:for-each>
            </tr>
        </xsl:if>

        <xsl:if test="Article[@addtr='3']">
            <tr>
                <xsl:for-each select="Article[@addtr='3']">

                    <xsl:apply-templates select="." />

                </xsl:for-each>
            </tr>
        </xsl:if>

        <xsl:if test="Article[@addtr='4']">
            <tr>
                <xsl:for-each select="Article[@addtr='4']">

                    <xsl:apply-templates select="." />

                </xsl:for-each>
            </tr>
        </xsl:if>

    </xsl:template>

    <xsl:template match="TopicList/Article">

        <td>
            <!--PIC START-->
            <xsl:if test="@IsPic ='Y'">
                <xsl:if test="xImgFile">
                    <img class="listimg">
                        <xsl:attribute name="src">
                            <xsl:value-of select="xImgFile" />
                        </xsl:attribute>
                        <xsl:attribute name="alt">
                            <xsl:value-of select="Caption" />
                        </xsl:attribute>
                    </img>
                </xsl:if>
            </xsl:if>
            <!--PIC END-->
            <!--TITLE start-->
            <p>
                <xsl:if test="xURL !=''">
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
                </xsl:if>
                <xsl:if test="xURL =''">
                    <xsl:value-of select="Caption" />
                </xsl:if>
                <!--TITLE END-->
                <br />
                <!--PostDate start-->
                <xsl:if test="@IsPostDate='Y'">
                    <xsl:value-of select="PostDate" />
                </xsl:if>
                <!--postdate end-->
            </p>
        </td>

    </xsl:template>
    <!-- 主資料區-->

    <!-- 分頁-->
    <xsl:template match="hpMain" mode="pageissue">
        <script>
            var totPage = <xsl:value-of select="totPage" />;
            totPage = totPage == 0 ? 1 : totPage;

            var qURL = "<xsl:value-of select="qURL" />";
            var perPage = <xsl:value-of select="PerPageSize" />
        </script>
        <div class="term">
            共 <xsl:value-of select="totRec" /> 筆資料，第 <xsl:value-of select="nowPage" />/<script>document.write(totPage)</script>頁，

            <xsl:if test="nowPage > 1">
                <a class="previous">
                    <xsl:attribute name="href">
                        lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage - 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
                    </xsl:attribute>
                    上一頁
                </a>
            </xsl:if>
            到第
            <select id="pickPage" onChange="pageChange()">
                <script>
                    <![CDATA[
	for (xi=1;xi<=totPage;xi++)
		document.write("<option value=" +xi + ">" + xi + "</option>");
	

	function pageChange() {
		goPage(pickPage.value);
	}
	function goPage(nPage) {
		document.location.href= "lp.asp?" + qURL + "&nowPage=" + nPage + "&pagesize=" + perPage 
	}
	function perPageChange() {
		document.location.href= "lp.asp?" + qURL + "&nowPage=" + pickPage.value + "&pagesize=" + perPage 
	}
]]>
                </script>
            </select>
            頁

            <xsl:if test="number(nowPage) &lt; number(totPage)">
                <a class="next">
                    <xsl:attribute name="href">
                        lp.asp?<xsl:value-of select="qURL" />&amp;nowPage=<xsl:value-of select="nowPage + 1" />&amp;pagesize=<xsl:value-of select="PerPageSize" />
                    </xsl:attribute>
                    下一頁
                </a>
            </xsl:if>
            <!--xsl:if test="queryYN"-->
            <div class="search">
                <A>
                    <xsl:attribute name="href">
                        qp.asp?<xsl:value-of select="xqURL" />
                    </xsl:attribute>
                    條件查詢
                </A>
            </div>
            <!--/xsl:if-->
            <script>
                pickPage.value = <xsl:value-of select="nowPage" />;
            </script>
        </div>
        <!--class="Page"-->
    </xsl:template>
    <!-- 分頁-->

</xsl:stylesheet>
