<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">
<!--會員 Start-->
<xsl:template match="hpMain" mode="xsMember">

	<xsl:if test="//login/status='true'">
		
      <p>農業知識入口網會員：<xsl:value-of select="//login/memName"/>[<xsl:value-of select="//login/memNickName"/>] [<xsl:value-of select="//login/memID"/>，您好<br/></p>
			<a href="/logout.asp" >登出會員</a>｜
			<a href="/sp.asp?xdURL=Coamember/member_Modify.asp&amp;mp=1" target="_blank" >修改個人資料</a>

		
</xsl:if>

<xsl:if test="//login/status='false'">
         
      <form method="post">
			<xsl:attribute name="action">/loginact.asp?<xsl:value-of select="//qStr"/></xsl:attribute>
      會員登入：
			<label>帳號
				<input name="account2" type="text" size="10" class="input" />
			</label>
			<label>密碼
				<input name="passwd2" type="password" size="10" class="input" />
			</label>			
			<input type="submit" name="submit" class="button" value="登入"/>
			<a href="/sp.asp?xdURL=coamember/member_Join.asp&amp;mp=1" target="_blank" >加入會員</a>｜
			<a href="/sp.asp?xdURL=coamember/member_Getpw.asp&amp;mp=1" target="_blank" >忘記密碼</a>
			</form>			
      
</xsl:if>
</xsl:template>
<!--會員 End-->
</xsl:stylesheet>
