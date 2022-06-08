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
<A href="genAPReport.asp">重GEN</A></TD></TR>
</TABLE>

<TABLE border="1" width="95%">
	<TR class="lightbluetable">
	<TH>類別</TH><TH>UseCase</TH><TH>HighLevel</TH><TH>Expanded</TH><TH>Contract</TH></TR>
	<xsl:apply-templates select="CatList/APCat" />
</TABLE>

</CENTER>

</BODY>
</HTML>
</xsl:template>


<xsl:template match="APCat">
	<TR><TD valign="TOP">	
		<xsl:attribute name="rowspan"><xsl:value-of select="count(UseCaseList/UseCase)" /></xsl:attribute>
		<xsl:value-of select="CatID" /> <BR/>
		<xsl:value-of select="CatName" />
		<xsl:apply-templates select="UseCaseList/UseCase" />	
	</TD></TR>
</xsl:template>

<xsl:template match="UseCase">
	<xsl:if test="position()=1">
	<TD class="whitetablebg">	
		<xsl:value-of select="APCode" /> :- 
		<xsl:value-of select="APName" />
	</TD>
	<TD class="whitetablebg"><xsl:apply-templates select="HighLevelSpec" /></TD>
	<xsl:apply-templates select="ExpandedSpec" />
	</xsl:if>
	<xsl:if test="position()!=1">
	<TR>
	<TD class="whitetablebg">	
		<xsl:value-of select="APCode" /> :- 
		<xsl:value-of select="APName" />
	</TD>
	<TD class="whitetablebg"><xsl:apply-templates select="HighLevelSpec" /></TD>
	<xsl:apply-templates select="ExpandedSpec" />
	</TR>
	</xsl:if>
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