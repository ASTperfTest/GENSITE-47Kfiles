<?xml version="1.0" encoding="utf-8" ?> 
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:user="urn:user-namespace-here" version="1.0">
  <xsl:output method="html" encoding="utf-8" omit-xml-declaration="yes" indent="yes" standalone="yes" /> 
  <xsl:template match="htPage">
     <html xmlns="http://www.w3.org/1999/xhtml" lang="zh-TW">
	<head>
		<title><xsl:value-of select="//pageSpec/pageHead" /></title>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<link href="css/form.css" rel="stylesheet" type="text/css" />
		<link href="css/layout.css" rel="stylesheet" type="text/css" />
		<script language="javascript">
			function checkOnSubmit(){
				<xsl:apply-templates select="//fieldList" mode="code" /> 
			}
		</script>
	</head>
	<body>
	 
		<div id="FuncName">
			<h1><xsl:value-of select="//pageSpec/pageHead" />/<xsl:value-of select="//pageSpec/pageFunction" /></h1>
			<div id="Nav">
				<xsl:apply-templates select="pageSpec/aidLinkList" /> 
				<A href="Javascript:window.history.back();" title="回前頁" >回前頁</A> 
			</div>
			<div id="ClearFloat"></div>
		</div>
		<div id="FormName"><xsl:value-of select="//pageSpec/pageHead" /></div>
	<form id="Form1" method="POST" name="reg"  action="AddAction.do" onSubmit="return checkOnSubmit()" >
		<table cellspacing="0">
     			<xsl:apply-templates select="htForm/formModel/fieldList/field" /> 
		</table>
		<input type="hidden" name="table">
			<xsl:attribute name="value"><xsl:value-of select="htForm/formModel/fieldList/tableName" /></xsl:attribute>
		</input>
            	<input type="submit" value="新增存檔" class="cbutton" />
                <input type="reset" class="cbutton" value="清除重填" />            
	  </form>
	  <div id="Explain">
		<h1>說明</h1>
		<ul>
			<li><span class="Must">*</span>為必要欄位</li>
		</ul>
	  </div>
	</body>
     </html>
  </xsl:template>
  <xsl:template match="field">
  	<tr>    
     		<td align="right" class="Label">
     		   <xsl:if test="inputType != 'hidden'" >
     			<xsl:if test="canNull='N'">
     				<span class="Must">*</span>
     			</xsl:if>
     			<xsl:value-of select="fieldLabel" />
     		   </xsl:if>
     		</td>     
      		<td class="whitetablebg">
      			<xsl:if test="inputType='textbox'">
      				<input type="text" class="InputText" >
      					<xsl:attribute name="name"> <xsl:value-of select="fieldName" /> </xsl:attribute>
      					<xsl:attribute name="size"> <xsl:value-of select="inputLen" /> </xsl:attribute>
      					<xsl:if test="initial != ''">
      						<xsl:attribute name="value"><xsl:value-of select="initial" /></xsl:attribute>
      					</xsl:if>
      				</input>
      			</xsl:if>
      			<xsl:if test="inputType='imgfile'">
      				<input type="file" class="InputText" >
      					<xsl:attribute name="name"> <xsl:value-of select="fieldName" /> </xsl:attribute>
      				</input>
      			</xsl:if>
      			<xsl:if test="inputType='hidden'">
      				<input>
      					<xsl:attribute name="type">hidden</xsl:attribute>
      					<xsl:attribute name="name"><xsl:value-of select="fieldName" /></xsl:attribute>
      					<xsl:if test="initial != ''">
      						<xsl:attribute name="value"><xsl:value-of select="initial" /></xsl:attribute>
      					</xsl:if>
      				</input>
      			</xsl:if>
      			<xsl:if test="inputType='refSelect'">
      				<select>
    					<xsl:attribute name="name"> <xsl:value-of select="fieldName" /> </xsl:attribute>
      					<xsl:apply-templates select="refLookup" mode="option" /> 
      				</select>
      			</xsl:if>
      		</td>     
     	</tr>
  </xsl:template>
  
  
  <xsl:template match="refLookup" mode="option">
    
  	<xsl:for-each select="code">
  		<xsl:choose>
  			<xsl:when test="@name = ../../initial">
  				<option  selected="selected">
  					<xsl:attribute name="value"> <xsl:value-of select="@name" /> </xsl:attribute>
  					<xsl:value-of select="@value" />
  				</option>
  			</xsl:when>
  			<xsl:otherwise>
  				<option>
  					<xsl:attribute name="value"> <xsl:value-of select="@name" /> </xsl:attribute>
  					<xsl:value-of select="@value" />
  				</option>
  			</xsl:otherwise>
  		</xsl:choose>
  	</xsl:for-each>
    
  </xsl:template>
  
  <xsl:template match="aidLinkList" >
  
  	   	<xsl:for-each select="Anchor">

  					<A>
  						<xsl:attribute name="href"><xsl:value-of select="AnchorURI" /></xsl:attribute>
  						<xsl:attribute name="title"><xsl:value-of select="AnchorDesc" /></xsl:attribute>
  						<xsl:value-of select="AnchorLabel" />
  					</A> 
  					
  		</xsl:for-each>
  	
  </xsl:template>
  
  <xsl:template match="refLookup" mode="option">
    
  	<xsl:for-each select="code">
  		<option>
  			<xsl:attribute name="value"> <xsl:value-of select="@name" /> </xsl:attribute>
  			<xsl:value-of select="@value" />
  		</option>
  	</xsl:for-each>
    
  </xsl:template>
  
   <xsl:template match="fieldList" mode="code" >
  	   	<xsl:for-each select="field">
  	   		<xsl:apply-templates select="." mode="code" /> 
  		</xsl:for-each>
  </xsl:template>
  
  <xsl:template match="field" mode="code" >
  		<xsl:if test="canNull = 'N'">
  			if( document.Form1.<xsl:value-of select="fieldName" />.value == '' ){
  				alert('請輸入 <xsl:value-of select="fieldLabel" />');
  				return false;
  			}
  		</xsl:if>
  </xsl:template>
  
</xsl:stylesheet>