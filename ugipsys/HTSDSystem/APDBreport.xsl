<?xml version="1.0"  encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="urn:user-namespace-here" version="1.0" >
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone= "yes"/>
	<xsl:template match="/">
		<HTML>
			<HEAD>
				<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
				<LINK rel="stylesheet" type="text/css" href="../inc/setstyle.css" />
			</HEAD>
		
<BODY text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<CENTER>
<TABLE border="0" width="90%">
<TR><TD align="right">
<A href="genAPDBReport.asp">重GEN</A></TD></TR>
</TABLE>
</CENTER>

	<xsl:apply-templates select="CatList/APCat" />

</BODY>
</HTML>
</xsl:template>

<xsl:template match="APCat">
	<BR/>
	<HR/>
	<BR/>
		系統代碼：<xsl:value-of select="CatID" /> <BR/>
		系統名稱：<xsl:value-of select="CatName" />

<TABLE border="1" width="95%">
	<TR class="lightbluetable">
	<TH>模 組</TH><TH>程 式 代 碼</TH><TH>程 式 名 稱</TH><TH>存 取 資 料 表</TH></TR>
	<xsl:apply-templates select="UseCaseList/UseCase" />	
</TABLE>

</xsl:template>

<xsl:template match="UseCaseList/UseCase">
	<TR><TD valign="TOP">	
		<xsl:attribute name="rowspan"><xsl:value-of select="count(ExpandedSpec/ContractSpec)" /></xsl:attribute>
		<xsl:value-of select="APCode" /> <BR/>
		<xsl:value-of select="APName" />
		<xsl:apply-templates select="ExpandedSpec/ContractSpec" />	
	</TD></TR>
</xsl:template>

<xsl:template match="ContractSpec">
	<xsl:if test="position()=1">
	<TD class="whitetablebg">	
		<xsl:value-of select="Name" /></TD> 
	<TD class="whitetablebg">	
		<xsl:value-of select="ResponsibilityList/Responsibility" />
	</TD>
	<TD class="whitetablebg"><TABLE border="0">
	<xsl:apply-templates select="htDEntityList/htDEntity" />
	</TABLE></TD>
	</xsl:if>
	<xsl:if test="position()!=1">
	<TR>
	<TD class="whitetablebg">	
		<xsl:value-of select="Name" /></TD>
	<TD class="whitetablebg">	
		<xsl:value-of select="ResponsibilityList//Responsibility" />
	</TD>
	<TD class="whitetablebg"><TABLE border="0">
	<xsl:apply-templates select="htDEntityList/htDEntity" />
	</TABLE></TD>
	</TR>
	</xsl:if>
</xsl:template>

<xsl:template match="htDEntity">
	<TR class="whitetablebg"><TD>
	<xsl:value-of select="." /></TD>
	<TD> </TD>
	<TD>
   <xsl:choose>
       <xsl:when test="@xRef='Y'">(R)</xsl:when>
      <xsl:otherwise>　</xsl:otherwise>
   </xsl:choose>
   </TD>
	<TD>
		<xsl:if test="@xUpdate='Y'">(W)
		</xsl:if>
   </TD>
   </TR>
</xsl:template>

<xsl:template match="HighLevelSpec">
	<span>
		<xsl:if test="Status='Y'">
			<xsl:attribute name="class">lightbluetable</xsl:attribute>
		</xsl:if>
		<A><xsl:attribute name="href">ucHighLevel.asp?apCode=<xsl:value-of select="../APCode" />&amp;verDate=<xsl:value-of select="Date" /> 
		</xsl:attribute>
		<xsl:value-of select="Date" /> 
		(<xsl:value-of select="Author" />)
		</A>
	</span>
	<BR/>
</xsl:template>

<xsl:template match="ExpandedSpec">
  <xsl:if test="position()=1">
	<TD class="whitetablebg">
	<span>
		<xsl:if test="Status='Y'">
			<xsl:attribute name="class">lightbluetable</xsl:attribute>
		</xsl:if>
		<A><xsl:attribute name="href">ucExpanded.asp?apCode=<xsl:value-of select="../APCode" />&amp;verDate=<xsl:value-of select="Date" /> 
		</xsl:attribute>
		<xsl:value-of select="Date" /> 
		(<xsl:value-of select="Author" />)
		</A>
	</span>
	<BR/>
	</TD>
	<TH class="whitetablebg"><xsl:apply-templates select="TypicalCourseOfEvents" /></TH>
  </xsl:if>
</xsl:template>

<xsl:template match="TypicalCourseOfEvents">
	<span>
		<xsl:if test="count(Event[@ContractXML]) = count(Event)">
			<xsl:attribute name="class">lightbluetable</xsl:attribute>
		</xsl:if>
	<xsl:value-of select="count(Event[@ContractXML])" /> / 
	<xsl:value-of select="count(Event)" />
	</span>
</xsl:template>


</xsl:stylesheet>