<%@ page language="java" contentType = "text/html;charset=UTF-8" %> 
<%@ taglib prefix="html" uri="/WEB-INF/tlds/struts-html.tld" %>
<%@ taglib prefix="logic" uri="/WEB-INF/tlds/struts-logic.tld" %>

<html:html xhtml="true">
<head>

<title>歡迎登入GIP管理系統</title>

<html:base />

<style type="text/css">
<!--
body {
	background-color: #161713;
	margin-left: 0%;
	margin-top: 50px;
	margin-right: 0%;
	margin-bottom: 50px;
}
-->
</style>
<link href="css/gip.css" rel="stylesheet" type="text/css" />

</head>

<body>

  <table cellspacing="0" ID="Table1">
  <tr><td><table cellspacing="0" ID="Table2">
     <tr>    
      <td align="right" class="Label"><span class="Must">*</span>目錄樹：</td>     
      <td class="whitetablebg">
            </td>     
     </tr>     
     <tr>    
      <td align="right" class="Label"><span class="Must">*</span>版面樣式：</td>     
      <td class="whitetablebg">
      </td></tr>
     </table></td>
  <td>
		<table cellspacing="0" border="1" ID="Table3">
		<tr align="center">
			<th>款式</th><th>代碼數值</th><th>說明事項</th>
		</tr>
		</table>
  <table cellspacing="0" border="1" ID="Table4">
  <tr align="center"><th>Tag</th><th>區塊名稱</th><th>目錄樹節點</th><th>則數</th></tr>
	</table></td></tr>
</TABLE>
<input type="button" value="編修存檔" name="Enter" class="cbutton" OnClick="formModSubmit()" ID="Button1"><% End IF %> 
   <input type="button" class=cbutton value="清除重填" onClick="resetForm()" ID="Button2" NAME="Button2">
</form>     

<div id="Explain">
	<h1>說明</h1>
	<ul>
		<li><span class="Must">*</span>為必要欄位</li>
	</ul>
</div>  		

</body>
</html:html>
