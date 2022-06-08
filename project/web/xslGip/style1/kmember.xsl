<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
<!--會員 Start-->
<xsl:template match="hpMain" mode="xsMember">

<xsl:if test="//login/status='true'">
<div class="login">
<h2>會員功能</h2>

<div class="body">
				
<p>農業知識入口網會員：<br/><xsl:value-of select="//login/memName"/>[<xsl:value-of select="//login/memNickName"/>] [<xsl:value-of select="//login/memID"/>]，您好。<br/></p>
<xsl:if test="//activity/active='true'">
	<p>目前您的知識一起答積分為：<xsl:value-of select="//activity/ActivityGrade"/>分</p>
</xsl:if>	
<br/>
<center>				
<xsl:if test="//webActivity/active='true'">
				<span class="buttonC">
					<a>
						<xsl:attribute name="href">/sp.asp?xdURL=webActivity/actPage1.asp&amp;mp=1&amp;id=<xsl:value-of select="//webActivity/id" /></xsl:attribute>
						參加網站使用性調查
					</a>
				</span>
			</xsl:if>		
			<span class="buttonC"><a href="/Member/MemberRanking.aspx?cat=A">會員積分排行榜</a></span>
			<!--span class="buttonC"><a href="/Planting/PlantingActIntro.aspx">虛擬種植活動首頁</a></span-->
<!--span class="buttonC"><a href="/sp.asp?xdURL=activity/qactivity-1.asp&amp;mp=1">參加滿意度調查</a></span-->
<span class="buttonB"><a href="logout.asp">登出會員</a></span>
<!--span class="buttonB"><a href="sp.asp?xdURL=Coamember/member_Getpw.asp&amp;mp=1">忘記密碼</a></span-->
<span class="buttonB"><a href="sp.asp?xdURL=Coamember/member_Modify.asp&amp;mp=1">修改個人資料</a></span>
<span class="buttonB"><a href="CatTree/CatTreeList.aspx?DatabaseId=DB020&amp;CategoryId=00A0101&amp;ActorType=002">消費者知識庫</a></span>
<span class="buttonB"><a href="CatTree/CatTreeList.aspx?DatabaseId=DB020&amp;CategoryId=00A0101&amp;ActorType=001">生產者知識庫</a></span>
	<xsl:if test="//login/gstyle='003'">
		<span class="buttonB"><a href="CatTree/CatTreeList.aspx?DatabaseId=DB020&amp;CategoryId=00A0101&amp;ActorType=003">學者知識庫</a></span>
	</xsl:if>

	<span class="buttonB">		
				<a>
					<xsl:attribute name="href">/knowledge/myknowledge_record.aspx</xsl:attribute>
					<xsl:attribute name="title">我的知識</xsl:attribute>
					我的知識
				</a>
			</span>
</center>
<div class="foot"></div>
</div>
			
</div>
</xsl:if>

<xsl:if test="//login/status='false'">
<div class="login">
<h2>會員登入</h2>
<div class="body">
<!--form method="post" action="sp.asp?xdURL=member/memberlogin_act.asp"-->
<form method="post" action="loginact.asp">
<label>帳號:
<input name="account2" type="text" size="20" class="txt" />
</label>
<label>密碼:
<input name="passwd2" type="password" size="20" class="txt" />
</label>
			
<input type="image" src="xslgip/style1/images3/login.gif" id="submit" name="submit"  class="btn" value="登入會員"/>				
</form>
			
<br/>

<center>
<span class="buttonB"><a href="sp.asp?xdURL=coamember/member_Join.asp&amp;mp=1">加入會員</a></span>
<span class="buttonB"><a href="sp.asp?xdURL=coamember/member_Getpw.asp&amp;mp=1">忘記密碼</a></span>
</center>			
</div>
<div class="foot"></div>
</div>
</xsl:if>
</xsl:template>
<!--會員 End-->
</xsl:stylesheet>
