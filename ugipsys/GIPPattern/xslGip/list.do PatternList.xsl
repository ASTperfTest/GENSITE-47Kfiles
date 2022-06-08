<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format">
  <xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" /> 

	<xsl:template match="/">
		<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
			
				<link rel="stylesheet" type="text/css">
					<xsl:attribute name="href">css/list.css</xsl:attribute>
				</link>
				<link rel="stylesheet" type="text/css">
					<xsl:attribute name="href">css/layout.css</xsl:attribute>
				</link>
				<link rel="stylesheet" type="text/css">
					<xsl:attribute name="href">css/setstyle.css</xsl:attribute>
				</link>
				
				<xsl:value-of select="//headScript"/>
			</head>
			<body>
				<div id="FuncName">
					<h1>
						<xsl:value-of select="//funcName"/>
					</h1>
					<div id="Nav">
						<xsl:for-each select="//navList/nav">
							<a>
								<xsl:attribute name="href"><xsl:value-of select="@href"/></xsl:attribute>
								<xsl:value-of select="."/>
							</a>
						</xsl:for-each>
						&amp;nbsp;
					</div>
				</div>
				<div id="FormName">
					<xsl:value-of select="//formName"/>
				</div>
				<xsl:choose>
					<xsl:when test="string-length(//mainContent/page/totaldata) > 0">
						<!-- page bar -->
						<xsl:apply-templates select="//mainContent/page"/>
						<!-- TopicList -->
						<xsl:choose>
							<xsl:when test="string-length(//mainContent/form/@name) > 0">
								<form>
									<xsl:attribute name="name"><xsl:value-of select="//mainContent/form/@name"/></xsl:attribute>
									<xsl:attribute name="method"><xsl:value-of select="//mainContent/form/@method"/></xsl:attribute>
									<xsl:attribute name="action"><xsl:value-of select="//mainContent/form/@action"/></xsl:attribute>
									<!-- UIParts -->
									<xsl:for-each select="//mainContent/UIParts/UIPart">
										<xsl:value-of select="."/>
									</xsl:for-each>
									<xsl:apply-templates select="//mainContent/TopicList"/>
								</form>
							</xsl:when>
							<xsl:otherwise>
								<!-- UIParts -->
								<xsl:for-each select="//mainContent/UIParts/UIPart">
									<xsl:value-of select="."/>
								</xsl:for-each>
								<xsl:apply-templates select="//mainContent/TopicList"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:when>
					<xsl:otherwise>
						<xsl:value-of select="//mainContent"/>
					</xsl:otherwise>
				</xsl:choose>
			</body>
		</html>
	</xsl:template>
	<xsl:template match="//mainContent/page">
		<div id="Page">
			<form name="list" action="list.do">
				<xsl:for-each select="//mainContent/page/hiddenInput/*">
					<input type="hidden">
						<xsl:attribute name="name"><xsl:value-of select="name(.)"/></xsl:attribute>
						<xsl:attribute name="value"><xsl:value-of select="."/></xsl:attribute>
					</input>
				</xsl:for-each>
  
    共
    <em>
					<xsl:value-of select="//mainContent/page/totaldata"/>
				</em>
    筆資料，每頁顯示
    <select name="pageSize" onchange="javascript:this.form.submit()">
					<option value="10">10</option>
					<option value="15">15</option>
					<option value="30">30</option>
					<option value="50">50</option>
					<option value="300">300</option>
				</select>
    筆，目前在第
    <select name="curPage" onchange="javascript:this.form.submit()">
					<script language="javascript">
						var totPage = <xsl:value-of select="//mainContent/page/totalPage"/>;
						
						<![CDATA[
							if (totPage==0){
								document.write("<option value=0 selected>0</option>");
							}else{
								for (xi=1;xi<=totPage;xi++){
									document.write("<option value=" +xi + ">" + xi + "</option>");
								}
							}
						
						]]></script>
				</select>
    頁
    <script language="javascript">
    list.pageSize.value = <xsl:value-of select="//mainContent/page/pageSize"/>;
    list.curPage.value = <xsl:value-of select="//mainContent/page/curPage"/>;
    </script>
			</form>
		</div>
	</xsl:template>
	<xsl:template match="//mainContent/TopicList">
		<table id="ListTable">
			<tbody>
				<tr>
					<xsl:for-each select="//mainContent/TopicList/ColumnHead/Column">
						<th>
							<xsl:choose>
								<xsl:when test="value = '[全選]'">
									<input type="button" value="全選" class="cbutton" name="ckall" onClick="Chkall"/>
								</xsl:when>
								<xsl:otherwise>
									<xsl:value-of select="value"/>
								</xsl:otherwise>
							</xsl:choose>
						</th>
					</xsl:for-each>
				</tr>
				<xsl:for-each select="//mainContent/TopicList/Group">
					<xsl:if test="@name != ''">
						<tr>
							<th align="left">
								<xsl:attribute name="colspan"><xsl:value-of select="count(//mainContent/TopicList/ColumnHead/Column)"/></xsl:attribute>
								<xsl:value-of select="@name"/>
							</th>
						</tr>
					</xsl:if>
					<xsl:for-each select="Article">
						<tr>
							<xsl:variable name="i" select="position()"/>
							<xsl:for-each select="Column">
								<td>
									<xsl:choose>
										<xsl:when test="value = '[checkbox]'">
											<input type="checkbox">
												<xsl:attribute name="name">ckbox<xsl:value-of select="$i"/></xsl:attribute>
											</input>
										</xsl:when>
										<xsl:when test="string(value) = '[button]'">
											<xsl:variable name="cid" select="@id"/>
											<input type="button" class="cbutton">
												<xsl:attribute name="value"><xsl:value-of select="//mainContent/TopicList/ColumnHead/Column[@id=$cid]/value"/></xsl:attribute>
												<xsl:attribute name="onclick">self.location='<xsl:value-of select="url"/>'</xsl:attribute>
											</input>
										</xsl:when>
										<xsl:when test="string(value) = '[radio]'">
											<input type="radio" name="ugrpID">
												<xsl:attribute name="value"><xsl:value-of select="url"/></xsl:attribute>
											</input>
										</xsl:when>
										<xsl:when test="string-length(url) > 0">
											<a>
												<xsl:attribute name="href"><xsl:value-of select="url"/></xsl:attribute>
												<xsl:attribute name="target"><xsl:value-of select="target"/></xsl:attribute>
												<xsl:value-of select="value"/>
											</a>
										</xsl:when>
										<xsl:otherwise>
											<xsl:value-of select="value"/>
										</xsl:otherwise>
									</xsl:choose>
								</td>
							</xsl:for-each>
						</tr>
					</xsl:for-each>
				</xsl:for-each>
			</tbody>
		</table>
		<xsl:for-each select="//mainContent/TopicList/FormInput/input">
			<xsl:copy-of select="."/>
		</xsl:for-each>
	</xsl:template>
</xsl:stylesheet>
