<?xml version="1.0"  encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="urn:user-namespace-here" version="1.0" >
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone= "yes"/>
	<xsl:template match="Version">
		<HTML>
			<HEAD>
				<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
				<LINK rel="stylesheet" type="text/css" href="../inc/setstyle.css" />
			</HEAD>
		
<BODY text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<CENTER>
<TABLE border="0" width="90%">
<TR><TD align="right">
<A><xsl:attribute name="HREF">JavaScript: history.back()</xsl:attribute>
Back</A></TD></TR>
</TABLE>
<TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="90%">
	<TR><TD class="lightbluetable" align="right" width="20%">AP類別-代碼/Operation</TD>
	  <TD class="whitetablebg">　
	  <xsl:value-of select="/OperationContract/APCat" />-
	  <xsl:value-of select="/OperationContract/Code" />
	  　 / 　<xsl:copy-of select="ContractSpec/Name" />  
	  (<xsl:copy-of select="ContractSpec/Type" />)
	  　　
	  <xsl:choose>
    	<xsl:when test="/OperationContract/formSpec">
			<A><xsl:attribute name="HREF">HT2CodeGen/uiView.asp?formID=
				<xsl:value-of select="/OperationContract/formSpec"/>&amp;progPath=
				<xsl:value-of select="/OperationContract/htCodePath"/></xsl:attribute>
				<xsl:attribute name="target">_UItry</xsl:attribute>
			[UI]</A>
			<A><xsl:attribute name="HREF"><xsl:value-of select="/OperationContract/htCodePath"/>/
			<xsl:value-of select="/OperationContract/formSpec"/>.asp</xsl:attribute>
			<xsl:attribute name="target">_blank</xsl:attribute>
			[run]
			</A>
    	</xsl:when>
	  </xsl:choose>
	  </TD></TR>
	<TR><TD class="lightbluetable" align="right" width="20%">版本</TD>
	  <TD class="whitetablebg">　
	  	<xsl:copy-of select="Date" /> -- 
	  	<xsl:copy-of select="Author" />
	  </TD></TR>
	<TR><TD class="lightbluetable" align="right" width="20%">Actors</TD>
	  <TD class="whitetablebg">　
	  	<xsl:apply-templates select=".//Actor" />
	  </TD></TR>
	<TR><TD class="lightbluetable" align="right" width="20%"><B>Responsibility</B></TD>
	  <TD class="whitetablebg">
		<xsl:apply-templates select="ContractSpec/ResponsibilityList/Responsibility" />
	  </TD></TR>
	<TR><TD class="lightbluetable" align="right" width="20%"><B>前置條件</B></TD>
	  <TD class="whitetablebg">
		<xsl:apply-templates select="ContractSpec/PreConditionList/PreCondition" />
	  </TD></TR>
	<TR><TD class="lightbluetable" align="right" width="20%"><B>動作連結</B></TD>
	  <TH class="whitetablebg">
	  	<TABLE border="1" width="80%">
			<xsl:apply-templates select="ContractSpec/AnchorList/Anchor" />
	  	</TABLE>
	  	</TH></TR>
	<TR><TD class="lightbluetable" align="right" width="20%"><B>說 明</B></TD>
	  <TD class="whitetablebg">
		<xsl:apply-templates select="ContractSpec/DescriptionList/Description" />
	  </TD></TR>
	<TR><TD class="lightbluetable" align="right" width="20%"><B>資料物件</B></TD>
	  <TD class="whitetablebg">
	  	<B>　參用：　</B>
			<xsl:apply-templates select="ContractSpec/RefObjectList/RefObject[@xRef='Y']" /><BR/>
	  	<B>　異動：　</B>
			<xsl:apply-templates select="ContractSpec/RefObjectList/RefObject[@xUpdate='Y']" /><BR/>
	  </TD></TR>
	<TR><TD class="lightbluetable" align="right" width="20%"><B>PostCondition</B></TD>
	  <TH class="whitetablebg">
	  	<TABLE border="1" width="80%">
			<xsl:apply-templates select="ContractSpec/PostConditionList/PostCondition" />
	  	</TABLE>
	  	</TH></TR>
</TABLE>
</CENTER>
<HR/>
	
	<xsl:apply-templates select="ReferenceDocumentList" />
	<HR/>
	<xsl:apply-templates select="TBDList" />	

</BODY>
</HTML>
</xsl:template>

	<xsl:template match="Actor">
		<xsl:value-of select="." />
		<xsl:if test="position()!=last()">, </xsl:if>
	</xsl:template>

	<xsl:template match="Responsibility">
		　<B><xsl:copy-of select="." /></B>
		<xsl:if test="position()!=last()"><BR/></xsl:if>
	</xsl:template>

	<xsl:template match="PreCondition">
		　<xsl:copy-of select="." />
		<xsl:if test="position()!=last()"><BR/></xsl:if>
	</xsl:template>

	<xsl:template match="Anchor">
		<TR class="whitetablebg">
			<TH class="lightbluetable"><xsl:value-of select="AnchorLabel" /></TH>
			<TD><xsl:value-of select="AnchorType" /></TD>
			<TD><xsl:value-of select="AnchorDesc" /></TD>
			<TD><xsl:value-of select="AnchorURI" /></TD>
		</TR>
	</xsl:template>

	<xsl:template match="Description">
		　<xsl:copy-of select="." />
		<xsl:if test="position()!=last()"><BR/></xsl:if>
	</xsl:template>

	<xsl:template match="RefObject">
		<xsl:value-of select="." />
		<xsl:if test="position()!=last()">, </xsl:if>
	</xsl:template>

	<xsl:template match="PostCondition">
		<TR class="whitetablebg">
			<TD class="lightbluetable" width="40%">　<xsl:value-of select="@eop" /></TD>
			<TD>　<xsl:copy-of select="." /></TD>
		</TR>
	</xsl:template>

	<xsl:template match="ReferenceDocumentList">
		<H4>參考文件</H4>
		<UL>
		<xsl:apply-templates select="ReferenceDocument" />
		</UL>
	</xsl:template>

	<xsl:template match="ReferenceDocument">
		<LI>
			<A target="_blank"><xsl:attribute name="HREF"><xsl:value-of select="URL" /></xsl:attribute>
	          <xsl:value-of disable-output-escaping="yes" select="DocName"/>
			</A><BR/>
			<xsl:copy-of select="Description" />
		</LI>
	</xsl:template>

	<xsl:template match="TBDList">
		<H4>待決定事項</H4>
		<UL>
		<xsl:apply-templates select="TBD" />
		</UL>
	</xsl:template>
	
	<xsl:template match="TBD">
		<LI>
			<Font size="2" color="red"><xsl:copy-of select="Subject" /></Font><br/>
			<xsl:copy-of select="Description" />
		</LI>
	</xsl:template>	

</xsl:stylesheet>