<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCode="HT011"
HTProgPrefix="htDField" %>
<!--#include virtual = "/inc/server.inc" -->
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="../inc/setstyle.css">
<title></title>
</head>
<!--#include virtual = "/inc/xdbutil.inc" -->
<%


	sqlCom = "SELECT * FROM HtDentity WHERE entityId=" & pkStr(request.queryString("entityID"),"")
	Set RSmaster = Conn.execute(sqlcom)


%>
<body>
<CENTER>
<form method="POST" name="reg" action="htDEntityCpSet.asp">
<CENTER><TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="80%">
<TR><TD class="lightbluetable" align="right">
Database 名稱：</TD>
<TD class="whitetablebg" colspan="3"><input name="htx_xDbName" size="30">
<input type="hidden" name="htx_entityID" value="<%=request.queryString("entityID")%>"> (本身db則不填)
</TD>
</TR>
<TR><TD class="lightbluetable" align="right"><font color="red">*</font>Table 名稱：</TD>
<TD class="whitetablebg" colspan="3"><input name="htx_xTableName" size="30" value="<%=RSmaster("tableName")%>">
</TD>
</TR>
</TABLE>

</CENTER>
<INPUT TYPE=submit name=doTask value="CopyIt">&nbsp;　&nbsp;
<INPUT TYPE=reset >
</form>     

<body>
</html>