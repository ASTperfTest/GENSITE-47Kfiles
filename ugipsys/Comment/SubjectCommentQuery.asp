<%@ CodePage = 65001 %>
<% 
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap = "單元資料維護"
HTProgFunc = "資料清單"
HTProgCode = "SComment"
HTProgPrefix = "SubjectComment" 
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
	<style type="text/css">
	<!--
		.style1 {color: #000000}
	-->
  </style>
</head>
<body>
	<form name="iForm" method="post" action="SubjectCommentList.asp">
	<INPUT TYPE="hidden" name="submitTask" value="Query" />	
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
	  <td width="50%" align="left" nowrap class="FormName">主題館意見查詢&nbsp;<font size=2></font></td>
		<td width="50%" class="FormLink" align="right"><A href="Javascript:window.history.back();" title="回前頁">回前頁</A></td>	
  </tr>
	<tr>
	  <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
	</tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
	<tr>
	  <td align=center colspan=2 width=80% height=230 valign=top>    
			<CENTER>
      <TABLE width="100%" id="ListTable">
			<TR>
				<Th>主題館：</Th>
				<TD class="eTableContent">
					<label>
						<input name="CtRootName" type="text" size="30" value="" />
					</label>
				</TD>
			</TR>
			<TR>
				<Th>文章標題：</Th>
				<TD class="eTableContent">
					<label>
						<input name="sTitle" type="text" size="30" value="" />
					</label>
				</TD>
			</TR>
			<TR>
				<Th>意見：</Th>
				<TD class="eTableContent">
					<input name="xBody" type="text" size="30" value="" />
				</TD>
			</TR>
			<TR>
				<Th>會員(帳號、姓名、暱稱)：</Th>
				<TD class="eTableContent">
					<input name="iEditor" type="text" size="30" value="" />
				</TD>
			</TR>
			</TABLE>
			</CENTER>
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center">     
					<input name="SubmitBtn" type="submit" class="cbutton" value="查詢">
					<input type="reset" value ="重　填" class="cbutton" ></td>
			</tr>
			</table>
    </td>
  </tr>  
	</table> 
	</form>
</body>
</html>
