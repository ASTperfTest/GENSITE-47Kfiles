<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>

<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<%
 sql="SELECT [fCTUPublic],[sTitle],[xImgFile],[xBody],[xImportant] FROM [mGIPcoanew].[dbo].[CuDTGeneric] where [iCUItem]="&request("iCUItem")

 set rs=conn.execute(sql)
 Set xup = Server.CreateObject("TABS.Upload")
 
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
	<meta name="GENERATOR" content="Hometown Code Generator 2.0">
	<meta name="ProgId" content="ePubEdit.asp">
	<title>編修表單</title>
	<link type="text/css" rel="stylesheet" href="../css/list.css">
	<link type="text/css" rel="stylesheet" href="../css/layout.css">
	<link type="text/css" rel="stylesheet" href="../css/setstyle.css">
    <script src="../Scripts/AC_ActiveX.js" type="text/javascript"></script>
    <script src="../Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
</head>




<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left">議題專區管理&nbsp;<font size=2>【編修專區】</td>
        <td width="50%" class="FormLink" align="right"><A href="Javascript:window.history.back();" title="回前頁">回前頁</A>	    </td>
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




<form method="POST" name="reg"  action="index_subject01.asp" ENCTYPE="MULTIPART/FORM-DATA">
<INPUT TYPE=hidden name="iCUItem" value="<%=request("iCUItem")%>">
<CENTER>
<TABLE  id="ListTable">
<TR>
  <Th><font color="red">*</font>
專區標題：</Th>
<TD class="eTableContent"><input name="sTitle" id="sTitle" size="50" value="<%=rs("sTitle")%>"></TD>
</TR>
<TR>
  <Th>專區簡介：</Th>
  <TD class="eTableContent"><textarea name="textarea" cols="50" rows="7"><%=rs("xBody")%></textarea></TD>
</TR>
<TR>
  <Th>專區圖檔：</Th>
  <TD class="eTableContent"><img src="/public/data/<%=rs("xImgFile")%>" />
	<input type="file" name="xImgFile">
  </TD>
</TR>
<!--TR>
  <Th>日期區間：</Th>
  <TD class="eTableContent"><input name="pcShowhtx_IDateS" size="8" readonly onClick="VBS: popCalendar 'htx_IDateS','htx_IDateE'" value="2008/1/1">
～
  <input name="htx_IDateS" type=hidden>
  <input name="htx_IDateE" type=hidden>
  <input name="pcShowhtx_IDateE" size="8" readonly onClick="VBS: popCalendar 'htx_IDateE',''" value="2008/7/31">
(選擇資料日期區間，可不設定)</TD>
</TR-->
<TR>
  <Th>是否公開：</Th>
  <TD class="eTableContent"><select name="select" id="select">
    <option value="Y" <%if  rs("fCTUPublic")="Y" then response.write "selected"%>>公開</option>
    <option value="N" <%if  rs("fCTUPublic")="N" then response.write "selected"%>>不公開</option>
  </select>  </TD>
</TR>
<TR>
  <Th>重要性：</Th>
  <TD class="eTableContent"><input name="htx_title22" size="10" value="<%=rs("xImportant")%>">
    (99為最高)</TD>
</TR>
</TABLE>
</CENTER>
<INPUT TYPE=hidden name=CalendarTarget>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
 <td align="center">     
        <input type="submit" value="編修存檔" class="cbutton" onClick="check()"  >
        <input type="button" value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
        <input type="reset" value ="重　填" class="cbutton" />
        <input type="button" value ="回前頁" class="cbutton" onClick="location.href='index.asp'">
 </td>
</tr>
</table>
</form>     

     
    </td>
  </tr>  
</table> 
</body>
</html>

