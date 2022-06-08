<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
    <!--xsl:include href="member.xsl" /-->
    <xsl:include href="hySearch.xsl"/>
    <xsl:template match="hpMain" mode="headinfo">
        <!-- header start -->
        <!--動態游標 Start-->
        <div id="cursor" style="height: 25px; border: left: 100px; position: absolute; z-index:1000; top: 160px; visibility: visible; width: 20px">
            <script language="JavaScript">
                Cursorpng = "";
                <xsl:if test="//SubjectBeaanCursor/CursorOpen='1'">
                    <xsl:if test="//SubjectBeaanCursor/CursorPic='1'">Cursorpng="springpng";</xsl:if>
                    <xsl:if test="//SubjectBeaanCursor/CursorPic='2'">Cursorpng="summerpng";</xsl:if>
                    <xsl:if test="//SubjectBeaanCursor/CursorPic='3'">Cursorpng="autumnpng";</xsl:if>
                    <xsl:if test="//SubjectBeaanCursor/CursorPic='4'">Cursorpng="winterpng";</xsl:if>
                    <xsl:if test="//SubjectBeaanCursor/CursorPic='5'">Cursorpng="autumnpng";</xsl:if>
                    <xsl:if test="//SubjectBeaanCursor/CursorPic='6'">Cursorpng="summerpng";</xsl:if>
                    <xsl:if test="//SubjectBeaanCursor/CursorPic='7'">Cursorpng="winterpng";</xsl:if>
                </xsl:if>
            </script>
            <img id="cursor-img" style="display: none" />
        </div>
		<script type="text/javascript" src="/aspSrc.asp" >/**/</script>
        <script src="js/mousecursor.js" type="text/javascript">/**/</script>
        <!--動態游標 End-->
        <!--div id="login" ><xsl:apply-templates select="." mode="xsMember" /></div-->
        <div id="header">
            <link rel="stylesheet" type="text/css">
                <xsl:attribute name="href">/css/pedia.css</xsl:attribute>
            </link>

            <link rel="stylesheet" type="text/css">
                <xsl:attribute name="href">/js/ui.datepicker.css</xsl:attribute>
                <xsl:attribute name="charset">utf-8</xsl:attribute>
            </link>
            <link rel="stylesheet" type="text/css">
                <xsl:attribute name="href">/css/jquery.tagsinput.css</xsl:attribute>
            </link>
            <link rel="stylesheet" type="text/css">
                <xsl:attribute name="href">/css/jquery.autocomplete.ForSubjectTags.css</xsl:attribute>
            </link>
            <link rel="stylesheet" type="text/css">
                <xsl:attribute name="href">/css/readHistory.css</xsl:attribute>
            </link>

            <script src="/js/pedia.js" language="javascript" >/**/</script>
            <script type="text/javascript" language="javascript" src="../js/jquery-1.5.2.min.js">/**/</script>
            <script src="/js/ui.datepicker.js" language="javascript">/**/</script>
            <script src="/js/ui.datepicker-zh-TW.js" language="javascript">/**/</script>
            <!--Modify by Max 　農業字典-->
            <script src="/js/jtip.js" type="text/javascript"></script>
            <style type="text/css" media="all">
                @import "/css/global.css";
                .style1
                {
                height: 23px;
                }
            </style>
            <!--End-->
            <!--Modify by Max 　站內單元圖片瀏覽器-->
            <script type="text/javascript" language="javascript" src="../js/yoxviewPic/yox.js">/**/</script>
            <script type="text/javascript" language="javascript" src="../js/yoxviewPic/yoxview-init.js">/**/</script>
            <script type="text/javascript" src="/js/jquery.autocomplete.js">/**/</script>
            <script type="text/javascript" src="/js/jquery.tagsinput.js">/**/</script>

            <script language="javascript" type="text/javascript">
                function NewAjax(id, url, Eventjs, dataArr) {
                $.ajax({
                type: 'post',
                url: url,
                dataType: "text",
                data: dataArr,
                cache: false,
                success: function(html) {
                if (id != "") {
                $("#" + id).html(html);
                }
                if (typeof Eventjs == "function") {
                Eventjs();
                }
                else if (typeof Eventjs == "string") {
                eval(Eventjs);
                }
                }
                });
                }

                $(document).ready(function () {
                $(".yoxview").yoxview({
                videoSize: { maxwidth: 720, maxheight: 560 },
                autoHideInfo:false,
                autoHideMenu:false,
                lang:'zh-tw'
                });

                });


            </script>
            <script type="text/javascript">
                var gaJsHost=(("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
                document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
            </script>
            <script type="text/javascript">
                try{
                var pageTracker=_gat._getTracker("UA-9195501-1");
                pageTracker._trackPageview();
                }catch(err)
                {}
            </script>
            <xsl:if test="//lockrightbtn='Y'">
                <script src="/nocopy.js" language="javascript" />
            </xsl:if>

            <!-- accesskey start -->
            <div class="accesskey">
                <a href="accesskey.htm" title="上方區塊" accesskey="U">:::</a>
            </div>
            <!-- accesskey start -->

            <span class="logo">
                <a>
                    <xsl:attribute name="href">
                        dp.asp?mp=<xsl:value-of select="mp" />
                    </xsl:attribute>
                    <img>
                        <xsl:attribute name="src">
                            <xsl:value-of select="TitlePicFilePath" />
                        </xsl:attribute>
                        <xsl:attribute name="alt">
                            <xsl:value-of select="rootname" />
                        </xsl:attribute>
                    </img>
                </a>
            </span>

            <ul>
                <li>
                    <xsl:if test="//login/status='true'">
                        <a>
                            <xsl:attribute name="href">/sp.asp?xdURL=Coamember/member_Modify.asp&amp;mp=1</xsl:attribute>
                            <xsl:value-of select="//login/memAccount" />
                        </a>
                        <a>
                            <xsl:attribute name="href">logout.asp</xsl:attribute>登出
                        </a>
                    </xsl:if>
                    <xsl:if test="//login/status='false'">
                        <a>
                            <xsl:attribute name="href">login.asp</xsl:attribute>登入
                        </a>
                    </xsl:if>
                </li>
                <li>
                    <a>
                        <xsl:attribute name="href">
                            dp.asp?mp=<xsl:value-of select="mp" />
                        </xsl:attribute>首頁
                    </a>
                </li>
                <!--li><a><xsl:attribute name="href">mailto:<xsl:value-of select="contactmail" /></xsl:attribute>聯絡我們</a></li-->
                <li>
                    <a>
                        <xsl:attribute name="href">
                            sitemap.asp?mp=<xsl:value-of select ="mp" />
                        </xsl:attribute>導覽
                    </a>
                </li>
                <li>
                    <a href="/SubjectList.aspx">
                        主題列表
                    </a>
                </li>
                <li>
                    <a href="/">
                        入口網
                    </a>
                </li>
            </ul>
            <xsl:apply-templates select="." mode="xsSearch" />

        </div>
        <!--橫幅圖片 start -->
        <div class="image">
            <img>
                <xsl:attribute name="src">
                    <xsl:value-of select="BannerPicFilePath" />
                </xsl:attribute>
                <xsl:attribute name="alt">
                    <xsl:value-of select="BannerPicFileName" />
                </xsl:attribute>
            </img>
        </div>
        <!--橫幅圖片 end-->
    </xsl:template>
</xsl:stylesheet>