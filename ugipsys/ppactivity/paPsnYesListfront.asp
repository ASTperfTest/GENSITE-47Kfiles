<% Response.Expires = 0
HTProgCap="���ʺ޲z"
HTProgFunc="���W�ǭ��M��"
HTProgCode="PA005"
HTProgPrefix="psEnroll" %>
<% response.expires = 0 %>

<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%

Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------


	sql = "SELECT htx.*, a.actName, a.actCat, a.actDesc, a.actTarget, xref2.mValue AS xrActCat" _
		& ", (SELECT count(*) FROM paEnroll AS e WHERE e.pasID=htx.pasID) AS erCount" _
		& " FROM paSession AS htx JOIN ppAct AS a ON a.actID=htx.actID" _
		& " LEFT JOIN CodeMain AS xref2 ON xref2.mCode = a.actCat AND xref2.codeMetaID='ppActCat'" _
		& " WHERE htx.paSID=" & session("paSID")
	set RS = conn.execute(sql)

%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<link rel="stylesheet" href="../inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td class="FormName"><%=RS("actName")%>&nbsp;<%=RS("dtNote")%>&nbsp;<font size=2>�i�����W��j</td>
    <td class="FormLink" valign="top" align=right>
	       <a href="Javascript:window.history.back();">�^�e��</a>
    </td>    
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height=230 valign=top>

    <CENTER>
     <TABLE width=95% cellspacing="1" cellpadding="6" class=bg border=0 bgcolor=336699>                   
     <tr align=center >    
	<td class=lightbluetable width=10% nowrap>�m�W</td>
	<td class=lightbluetable >Email</td>
	<td class=lightbluetable>�ͤ�</td>
	<td class=lightbluetable width=20%>���</td>
	<td class=lightbluetable align=left>��}</td>
   </tr>	                                
<%                   
	fSql = "SELECT htx.paSID, htx.ckValue, htx.erDate, htx.status, htx.psnID AS xPsnID, p.*" _
		& " FROM (paEnroll AS htx LEFT JOIN paPsnInfo AS p ON p.psnID = htx.psnID)" _
		& " WHERE htx.status='Y' AND htx.pasID=" & session("paSID")
	set RSreg = conn.execute(fSql)

    while not RSreg.eof           
%>                  
<tr>                  
	<TD class=whitetablebg align=center><font size=2><%=RSreg("pName")%></font></td>
	<TD class=whitetablebg align=center><font size=2><%=RSreg("email")%></font></td>
	<TD class=whitetablebg align=center><font size=2><%=d7date(RSreg("birthDay"))%></font></td>
	<TD class=whitetablebg align=center><font size=2><%=RSreg("myOrg")%></font></td>
	<TD class=whitetablebg align=left><font size=2><%=RSreg("zipcode")%>�@<%=RSreg("addr")%></font></td>
	

    </tr>
    <%
         RSreg.moveNext
	wend
   %>
    </TABLE>
    </CENTER>
       <!-- �{������ ---------------------------------->  
     </form>  
    </td>
  </tr>  
</table> 
</body>
</html>                                 
