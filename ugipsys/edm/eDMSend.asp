<%@ CodePage = 65001 %>
<% 
	Response.Expires = 0
	response.charset="utf-8"
	HTProgCode = "DATAEDM"

	Dim baseDSDId : baseDSDId = "7"
	Dim ctUnitId : ctUnitId = "2162"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%	
	
	sql = "SELECT TOP 1 * FROM CuDTGeneric WHERE iCTUnit = " & ctUnitId & " ORDER BY xPostDate DESC"
	Set rs = conn.execute(sql)
	If rs.eof Then
		response.write "<script>alert('找不到設定值');history.back();</script>"
	Else
		Dim sTitle : sTitle = ""
		Dim xBody : xBody = ""
		if not rs.eof then
			sTitle = rs("sTitle")
			xBody = rs("xBody")
		end if
		Set rs = nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<title>編修表單</title>
	<link type="text/css" rel="stylesheet" href="/css/list.css">
	<link type="text/css" rel="stylesheet" href="/css/layout.css">
	<link type="text/css" rel="stylesheet" href="/css/setstyle.css">    
</head>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td width="50%" class="FormName" align="left">eDM管理&nbsp;<font size=2>【編修】</td>
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
		<form method="POST" name="reg" action="eDMSendAct.asp">
		<INPUT TYPE="hidden" name="submitTask" value="">
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td align="center"><input type="submit" value="寄送eDM" name="button4"  class="cbutton" ></td></tr>
		</table>
		<CENTER>
		<TABLE  id="ListTable">
		<TR>
			<Th>eDM標題：</Th>
			<TD class="eTableContent"><%=sTitle%></TD>
		</TR>
		<TR>
			<Th>eDM內容：</Th>
			<TD class="eTableContent">
				<%=xBody%>
			</TD>
		</TR>
		</TABLE>
		</CENTER>				
		</form>     
		</td>
	</tr>
	</table>
</body>
</html>
<%
	end if 	
%>

