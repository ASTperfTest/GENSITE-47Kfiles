<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
<!--Copyright Start-->
	<xsl:template match="hpMain" mode="Copyright">
		<!-- footer start -->
		<div id="footer">
			<!--a href="mp.asp"><img src="xslgip/styleB/images/aplus.jpg" alt="通過 A+ 優先等級無障礙網頁檢測" width="86" height="29" /></a-->
			<img src="xslgip/images_data/60x80_best.gif"/>
            <p>
                本網站為行政院農業委員會版權所有，所刊載之內容均受著作權法保護，歡迎連結使用。禁止未經授權之複製或下載等用於營利行為，違者依法必究。<br/> 行政院農業委員會 版權所有 &amp;copy; 2007 All Rights Reserved.<![CDATA[　]]>
                <xsl:choose>
                    <xsl:when test="footer_dept=''">
                    </xsl:when>
                    <xsl:otherwise>
                        <script type="text/javascript">
                            var vc1 = '<xsl:value-of select="CounterAll" />'
                            var vc2 = '<xsl:value-of select="CounterThisYear" />'
                            var vc3 = '<xsl:value-of select="CounterThisMonth" />'
                        </script>
                        <span onclick='alert("歷年來瀏覽人數:"+ vc1 +" \r\n本年度瀏覽人數:"+ vc2 +" \r\n本月份瀏覽人數:" + vc3);'>
                            維護單位：
                        </span>
                        <a target='_blank'>
                            <xsl:attribute name="href">
                                <xsl:value-of select="footer_dept_url" />
                            </xsl:attribute>
                            <xsl:value-of select="footer_dept" />
                        </a>
                        <br />
                    </xsl:otherwise>
                </xsl:choose>
                最佳瀏覽狀態為 IE7.0 以上, 1024*768 解析度
            </p>
		</div>
		<br />
<!-- footer end -->
	</xsl:template>
<!--Copyright End-->	

</xsl:stylesheet>


