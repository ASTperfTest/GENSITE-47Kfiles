﻿<?xml version="1.0"  encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="urn:user-namespace-here" version="1.0" >
<xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone= "yes"/>
	<xsl:template match="dsTable">
		<HTML>
			<HEAD>
				<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
				<LINK rel="stylesheet" type="text/css" href="../inc/setstyle.css" />
			</HEAD>
		
<BODY text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<CENTER>
　<BR/>
<TABLE border="1" class="bluetable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="whitetablebg" colspan="6" style="font:bold 150% 標楷體;">
Table Name: <xsl:value-of select="tableName" />
　( <xsl:value-of select="tableDesc" /> )
</TD>
<TD class="whitetablebg" align="right">
<xsl:value-of select="reportDate" />
</TD>
</TR>
<TR class="lightbluetable">
<TH>欄位名稱</TH>
<TH>標題</TH>
<TH>資料型別</TH>
<TH>長度</TH>
<TH>鍵值</TH>
<TH>Null</TH>
<TH>參照代碼</TH>
</TR>
	<xsl:apply-templates select="fieldList/field" />


</TABLE>
</CENTER>
</BODY>
</HTML>
</xsl:template>

	<xsl:template match="field">
		<TR class="whitetablebg">
			<TD><xsl:copy-of select="fieldName" /></TD>
			<TD><xsl:copy-of select="fieldLabel" /></TD>
			<TD><xsl:copy-of select="dataType" /></TD>
			<TD align="right"><xsl:copy-of select="dataLen" /></TD>
			<TD align="center">	  
				<xsl:if test="isPrimaryKey='Y'">
					<xsl:copy-of select="isPrimaryKey" />
				</xsl:if>
			</TD>
			<TD align="center">	  
				<xsl:if test="canNull='Y'">
					<xsl:copy-of select="canNull" />
				</xsl:if>
			</TD>
			<TD><xsl:copy-of select="refLookup" /></TD>
		</TR>
	</xsl:template>

	<xsl:template match="xdsTable">
		<HTML>
			<HEAD>
				<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
				<LINK rel="stylesheet" type="text/css" href="../inc/setstyle.css" />
			</HEAD>
		
<BODY text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<CENTER>
<TABLE border="1" width="90%">
<TR><TD align="right">
<A><xsl:attribute name="HREF">ucVersion.asp?apCode=<xsl:value-of select="/UseCase/Code" /></xsl:attribute>
Back</A></TD></TR>
</TABLE>
<TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="90%">
	<TR><TD class="lightbluetable" align="right" width="20%">類別/AP代碼</TD>
	  <TD class="whitetablebg">　
	  <xsl:copy-of select="/UseCase/APCat" /> -
	  <xsl:copy-of select="/UseCase/Code" />
	  </TD></TR>
	<TR><TD class="lightbluetable" align="right" width="20%">模組名稱</TD>
	  <TD class="whitetablebg">　
	  <xsl:copy-of select="/UseCase/Name" />
	  </TD></TR>
	<TR><TD class="lightbluetable" align="right" width="20%">版本</TD>
	  <TD class="whitetablebg">　
	  	<xsl:copy-of select="Date" /> (<xsl:copy-of select="Kind" />) --
	  	<xsl:copy-of select="Author" />
	  </TD></TR>
	<TR><TD class="lightbluetable" align="right" width="20%">Actors</TD>
	  <TD class="whitetablebg">　
	  	<xsl:apply-templates select=".//Actor" />
	  </TD></TR>
	<TR><TD class="lightbluetable" align="right" width="20%">Purpose</TD>
	  <TD class="whitetablebg">　
	  	<xsl:apply-templates select=".//Purpose" />
	  </TD></TR>
</TABLE>
</CENTER>
<HR/>
	
	<xsl:apply-templates select="ExpandedSpec/PreCondition" />
	<HR/>
	<xsl:apply-templates select="ExpandedSpec/DescriptionList" />
	<HR/>
	<xsl:apply-templates select="ExpandedSpec/TypicalCourseOfEvents" />
	<HR/>
	<xsl:apply-templates select="ReferenceDocumentList" />
	<HR/>
	<xsl:apply-templates select="TBDList" />	

</BODY>
</HTML>
</xsl:template>

	<xsl:template match="pageTitle">
		<TITLE>
			<xsl:value-of select="." />
		</TITLE>
	</xsl:template>

	<xsl:template match="Actor">
			<xsl:value-of select="." />
		<xsl:if test="@initiator='Y'">
			 (Initiator)
		</xsl:if>
		<BR/>
	</xsl:template>

	<xsl:template match="PreCondition">
		<B>前置條件： </B>
			<xsl:copy-of select="." />
	</xsl:template>

	<xsl:template match="DescriptionList">
		<H2>規格描述</H2>
		<BLOCKQUOTE>
		<xsl:apply-templates select="Description" />
		</BLOCKQUOTE>
	</xsl:template>

	<xsl:template match="Description">
		<P>
			<xsl:copy-of select="." />
		</P>
	</xsl:template>

	<xsl:template match="TypicalCourseOfEvents">
		<H2>TypicalCourseOfEvents</H2>
		<BLOCKQUOTE>
		<TABLE border="1" width="90%">
		<TR><TH>Client</TH><TH>Server</TH></TR>
		<xsl:apply-templates select="Event" />
		</TABLE>
		</BLOCKQUOTE>
	</xsl:template>

	<xsl:template match="Event">
	  <xsl:if test="@type='Client'">
		<TR><TD>
			<xsl:copy-of select="." />
			<xsl:choose>
				<xsl:when test="@ContractXML">
					<A><xsl:attribute name="HREF">viewContract.asp?xmlFile=
						<xsl:value-of select="@ContractXML"/></xsl:attribute>
					[<xsl:value-of select="@ContractXML"/>]</A>
				</xsl:when>
      			<xsl:otherwise>
					<BUTTON>
					<xsl:attribute name="onClick">VBS: getXMLName '<xsl:value-of select="@id"/>'</xsl:attribute>
					新Contract</BUTTON>
      			</xsl:otherwise>
			</xsl:choose>
   		</TD><TD>　</TD></TR>
	  </xsl:if>
	  <xsl:if test="@type='Server'">
		<TR><TD>　</TD><TD>
			<xsl:copy-of select="." />
		</TD></TR>
	  </xsl:if>
	</xsl:template>

	<xsl:template match="ReferenceDocumentList">
		<H2>參考文件</H2>
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
		<H2>待決定事項</H2>
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