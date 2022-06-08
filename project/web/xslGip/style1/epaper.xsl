<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:user="urn:user-namespace-here" version="1.0">

<!-- 電子報 Start-->
<xsl:template match="hpMain" mode="xsEPaper">
	<div class="Search">
		<div class="Head">訂閱電子報</div>
		<div class="body">
		<form method="post" action="member/epaper_act.asp?CtRootID=33">
			<label>E-mail: </label>
				<input name="email" type="text" id="input_area" value="" size="13"/>
				<br />
				<input name="submit1" type="image" src="xslGip/style1/images/subscribe.gif" alt="確定訂閱電子報" border="0"
					align="Middle" />
				<input name="submit2" type="image" src="xslGip/style1/images/cancel.gif" alt="取消訂閱電子報" border="0"
					align="middle" />
			</form>
		</div>
	</div>
</xsl:template>
<!-- 電子報 End-->	

</xsl:stylesheet>


