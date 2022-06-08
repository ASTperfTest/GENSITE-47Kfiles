<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	HTProgCap="單元資料維護"
	HTProgFunc="編修"
	HTUploadPath=session("Public")+"data/"
	HTProgCode="CW02"
	HTProgPrefix="CW02" 
%>
<!--#include virtual = "/inc/server.inc" -->
<!-- #INCLUDE FILE="../inc/dbFunc.inc" -->
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	<link rel="stylesheet" href="/inc/setstyle.css">
	<link type="text/css" rel="stylesheet" href="/css/list.css">
	<link type="text/css" rel="stylesheet" href="/css/layout.css">
	<link type="text/css" rel="stylesheet" href="/css/setstyle.css">
<title></title>
</head>
<body>
	<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
    <td width="50%" align="left" nowrap class="FormName">會員KPI排行榜設定&nbsp;</td>
		<td width="50%" class="FormLink" align="right"></td>	
  </tr>
	<tr>
	  <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
	</tr>
	<tr>
	  <td class="Formtext" colspan="2" height="15"></td>
  </tr>  
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
		<Form name=reg method="POST">  
    <CENTER>
		<table width=100% cellpadding="0" cellspacing="1" id="ListTable">
    <tr align=left>
			<th>排行榜名稱</th>
      <th>詳細設定</th>
    </tr>
    <tr>
			<td class=eTableContent><a href="KpiBasicSetting.asp?id=st_1">會員積分總排行</a></td>
      <td class=eTableContent><a href="KpiMemberRankSetting.asp?id=st_1&phase=edit">權重設定</a></td>
    </tr>
    <tr>
      <td class=eTableContent><a href="KpiBasicSetting.asp?id=st_2">知識吸收度</a></td>
      <td class=eTableContent><a href="KpiBrowseSetting.asp?id=st_2&phase=edit">給分設定</a></td>
    </tr>
    <tr>
      <td class=eTableContent><a href="KpiBasicSetting.asp?id=st_3">知識持久度</a></td>
      <td class=eTableContent><a href="KpiLoginSetting.asp?id=st_3&phase=edit">給分設定</a></td>
    </tr>
    <tr>
      <td class=eTableContent><a href="KpiBasicSetting.asp?id=st_4">知識分享度</a></td>
      <td class=eTableContent><a href="KpiShareSetting.asp?id=st_4&phase=edit">給分設定</a></td>
    </tr>
    <tr>
      <td class=eTableContent><a href="KpiBasicSetting.asp?id=st_5">知識貢獻度</a></td>
      <td class=eTableContent><a href="KpiGradeSetting.asp?id=st_5&phase=edit">給分設定</a></td>
    </tr>
    <tr>
      <td class=eTableContent><a href="KpiBasicSetting.asp?id=st_6">知識踴躍度</a></td>
      <td class=eTableContent><a href="KpiAdditionalSetting.asp?id=st_6&phase=edit">給分設定</a></td>
    </tr>
    </table>
    </CENTER>    
		<table border="0" width="100%" cellspacing="0" cellpadding="0"></table>
    </form>  
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center"></td>
	</tr>
</table> 
</body>
</html>                                 
