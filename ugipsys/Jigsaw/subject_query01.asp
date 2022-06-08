<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>
<%
	'response.write session("jigsqlindex")
	'response.write session("jigsqlindex1")
  session("jigsqlindex")=""
  session("jigsqlindex1")=""

%>
<!--#include virtual = "/inc/server.inc" -->
<html> 
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <meta name="ProgId" content="kpiQuery.asp">
    <title>查詢表單</title>
	<link type="text/css" rel="stylesheet" href="../css/list.css">
	<link type="text/css" rel="stylesheet" href="../css/layout.css">
	<link type="text/css" rel="stylesheet" href="../css/setstyle.css">
    <script src="../Scripts/AC_ActiveX.js" type="text/javascript"></script>
    <script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
    <script language="javascript" src="js/SS_popup.js" type="text/javascript"></script>
	<style type="text/css">
<!--
.style1 {color: #000000}
-->
    </style>
</head>
<body>
	<form name="iForm" method="post" action="index.asp">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
	    <td width="50%" align="left" nowrap class="FormName">農業推薦單元知識拼圖內容管理&nbsp;
		<font size=2>【內容條例清單--主題專區 專區查詢】</font>
		</td>
		<td width="50%" class="FormLink" align="right">
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A>			
		</td>	
  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    
		<CENTER>
        <TABLE width="100%" id="ListTable">
  <TR>
  <Th>資料標題：</Th>
  <TD class="eTableContent"><label>
    <input name="sTitle" id = "sTitle" size="30">
  (輸入資料標題)</label></TD>
</TR>
<TR>
  <Th>日期範圍：</Th>
  <TD class="eTableContent">
  <input type="text" name="value(startDate)" id="startDate" size="10" value="">
							<img
								src="icon_cal.gif"
								alt="日曆" style="cursor: hand" border="0"
								onclick="fPopUpCalendarDlgFormat(document.forms[0]['value(startDate)'],0);" />
								
	~
							<input type="text" name="value(endDate)" size="10" value="">
							<img
								src="icon_cal.gif"
								alt="日曆" style="cursor: hand" border="0"
								onclick="fPopUpCalendarDlgFormat(document.forms[0]['value(endDate)'],0);" />		
      (選擇資料日期區間)</TD>
</TR>
<TR>
  <Th>狀態：</Th>
  <TD class="eTableContent"><span class="FormLink">
    <Select name="Status" size=1>
      <option value="">請選擇</option>
      <option value="Y" selected>公開</option>
      <option value="N">不公開</option>
    </select>
    <span class="style1">(預設為公開)</span></span></TD>
</TR>

</TABLE>
</CENTER>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
 <td align="center">     
        <input name="button4" type="submit" class="cbutton" value="查詢">
        <input type="reset" value ="重　填" class="cbutton" ></td>
</tr>
</table>

	

     
    </td>
  </tr>  
</table> 
</form>
</body>
</html>
