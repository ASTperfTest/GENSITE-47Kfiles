<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
    <!--會員 Start-->
    <xsl:template match="hpMain" mode="xsMember">

        <xsl:if test="//login/status='true'">
            <div class="login2">
                <h2>會員功能</h2>
                <div class="body">
                    <p>
                        農業知識入口網會員：<br/>
                        <xsl:if test="//login/memID != ''">
                            <xsl:value-of select="//login/memID"/>
                        </xsl:if>

                        <xsl:choose>
                            <xsl:when test="//login/memNickName != ''">
                                [<xsl:value-of select="//login/memNickName"/>]
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:if test="//login/memName != ''">
                                    <xsl:choose>
                                        <xsl:when test="string-length(//login/memName) &lt;=2 ">
                                            [<xsl:value-of select="substring(//login/memName,1,1) "/>*]
                                        </xsl:when>
                                        <xsl:otherwise>
                                            [<xsl:value-of select="substring(//login/memName,1,1) "/>*<xsl:value-of select="substring(//login/memName, string-length(//login/memName),1) "/>]
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </xsl:if>
                            </xsl:otherwise>
                        </xsl:choose>
                        , 您好。
                    </p>
                    <xsl:if test="//activity/active='true'">
                        <p>
                            目前您的知識一起答積分為：<xsl:value-of select="//activity/ActivityGrade"/>分
                        </p>
                    </xsl:if>
                    <ul>
                        <xsl:if test="//webActivity/active='true'">
                            <li class="activity">
                                <a>
                                    <xsl:attribute name="href">/kmactivity/kmwebpuzzle/06.aspx?a=puzzle</xsl:attribute>
                                   愛拼才會贏個人專區
                                </a>
                            </li>
                        </xsl:if>
						<xsl:if test="//webActivity/active='true'">
						<li class="activity"><a>
						<xsl:attribute name="href">/sp.asp?xdURL=webActivity/actPage1.asp&amp;mp=1&amp;id=<xsl:value-of select="//webActivity/id" /></xsl:attribute>
						參加網站使用性調查
						</a>
						</li>
					</xsl:if>
                        <li class="em">
                            <a href="/Member/MemberRanking.aspx?cat=A">會員積分排行榜</a>
                        </li>
                        <li class="em">
                            <a href="/Member/MemberRankingThisYear.aspx?cat=A">會員積分排行榜(本年度)</a>
                        </li>
                        <li>
                            <a>
                                <xsl:attribute name="href">
                                    /logout.asp?redirecturl=<xsl:value-of select="//redirectUrl"/>
                                </xsl:attribute>
                                登出會員
                            </a>
                        </li>
                        <li>
                            <a href="/sp.asp?xdURL=Coamember/member_Modify.asp&amp;mp=1">修改個人資料</a>
                        </li>
                        <li>
                            <a href="/CatTree/CatTreeList.aspx?DatabaseId=DB020&amp;CategoryId=00A0101&amp;ActorType=002">消費者知識庫</a>
                        </li>
                        <xsl:if test="//login/gstyle='003' or //login/gstyle='005'">
                            <li>
                                <a href="/CatTree/CatTreeList.aspx?DatabaseId=DB020&amp;CategoryId=00A0101&amp;ActorType=003">學者知識庫</a>
                            </li>
                        </xsl:if>
                        <li>
                            <a>
                                <xsl:attribute name="href">/knowledge/myknowledge_record.aspx</xsl:attribute>
                                <xsl:attribute name="title">我的知識</xsl:attribute>
                                我的知識
                            </a>
                        </li>
                        <!-- Added By Leo   2011-06-23    好文推薦   Start   -->
                        <li>
                            <a>
                                <xsl:attribute name="href">/recommand/recommand_Add.aspx</xsl:attribute>
                                <xsl:attribute name="title">推薦好文</xsl:attribute>推薦好文
                            </a>
                        </li>
                        <!-- Added By Leo   2011-06-23    好文推薦    End    -->
                        <!-- Added By Leo   2011-06-23    好友推薦   Start   -->
                        <li>
                            <a>
                                <xsl:attribute name="href">/Member/MemberInvitePage.aspx?</xsl:attribute>
                                <xsl:attribute name="title">邀請朋友</xsl:attribute>
                                邀請朋友
                            </a>
                        </li>
                        <!-- Added By Leo   2011-06-23    好友推薦    End    -->                        
                    </ul>
                </div>
            </div>
        </xsl:if>

        <xsl:if test="//login/status='false'">
            <div class="login">
                <h2>會員登入</h2>
                <div class="body">
                    <form method="post">
                        <xsl:attribute name="action">
                            /loginact.asp?redirecturl=<xsl:value-of select="//redirectUrl"/>
                        </xsl:attribute>
                        <label>
                            帳號:
                            <input name="account2" id="id_account2" type="text" size="20" class="txt" />
                        </label>
                        <label>
                            密碼:
                            <input name="passwd2" id="id_passwd2" type="password" size="20" class="txt" />
                        </label>
                        <center>
                            <input type="image" id="submit" name="submit"  class="btn" value="登入會員" OnClick="javascript:xsubmit()" >
                                <xsl:attribute name="src">
                                    /xslgip/<xsl:value-of select="myStyle"/>/images/login.gif
                                </xsl:attribute>
                            </input>
                        </center>
                    </form>
                    <ul>
                        <li>
                            <a href="/sp.asp?xdURL=coamember/member_Join.asp&amp;mp=1">加入會員</a>
                        </li>
                        <li>
                            <a href="/sp.asp?xdURL=coamember/member_Getpw.asp&amp;mp=1">忘記密碼</a>
                        </li>
                    </ul>
                </div>
            </div>
        </xsl:if>
        <script language="javascript">
            <![CDATA[
	
	function doEncodeField(fieldName){
		var newString = doEncodeString(document.getElementById(fieldName).value);    
		return newString;
	}
	function doEncodeString(inputString){
		var len=inputString.length;
		var publicKey=18;
		var rs="";
		for(var i=0;i<len;i++){
			rs+=String.fromCharCode(inputString.charCodeAt(i)+publicKey);
		}
		return rs;
	}
	
	function xsubmit() 
	{
		if( document.getElementById("id_account2").value == "" ) {
			alert("帳號未填");
			document.getElementById("id_account2").focus();
			window.event.returnValue = false;
		}
		else if( document.getElementById("id_passwd2").value == "" ) {
			alert("密碼未填");
			document.getElementById("id_passwd2").focus();
			window.event.returnValue = false;
		}
		else {
			//document.getElementById("id_account2").value = doEncodeField("id_account2");
			//document.getElementById("id_passwd2").value = doEncodeField("id_passwd2");
			//document.getElementById("f1").submit();
		}
	}
	]]>
        </script>
    </xsl:template>
    <!--會員 End-->
    <xsl:template match="hpMain" mode="newest">
        <xsl:if test = "//newest/article">
            <div class="rbox">
                <h2>最近瀏覽</h2>
                <div class="body">
                    <ul>
                        <xsl:for-each select="//newest/article">
                            <li>
                                <a>
                                    <xsl:attribute name="href">
                                        <xsl:value-of select="url" />
                                    </xsl:attribute>
                                    <xsl:value-of select="title" />
                                </a>
                            </li>
                        </xsl:for-each>
                    </ul>
                    <!--div class="float"><a href="#">清除記錄</a></div-->
                    <!--位置移到body內 -->
                </div>
                <div class="Foot"></div>
            </div>
        </xsl:if>
    </xsl:template>
</xsl:stylesheet>
